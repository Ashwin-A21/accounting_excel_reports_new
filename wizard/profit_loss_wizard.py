from odoo import models, fields, api
from odoo.exceptions import UserError
import xlsxwriter
from io import BytesIO
import base64
from collections import defaultdict

class ProfitLossWizard(models.TransientModel):
    _name = 'profit.loss.wizard'
    _description = 'Profit & Loss Report Wizard'

    start_date = fields.Date(string='Start Date', required=True)
    end_date = fields.Date(string='End Date', required=True)
    company_id = fields.Many2one('res.company', string='Company', 
                                 required=True, 
                                 default=lambda self: self.env.company)
    excel_file = fields.Binary(string='Excel File', readonly=True)
    file_name = fields.Char(string='File Name', readonly=True)
    line_ids = fields.One2many('tally.profit.loss.line', 'wizard_id', string='Report Lines')

    @api.constrains('start_date', 'end_date')
    def _check_dates(self):
        for record in self:
            if record.start_date > record.end_date:
                raise UserError('End Date must be greater than Start Date!')

    def _classify_pl_account_to_tally_group(self, account):
        """Standalone Tally Classification for P&L"""
        name = (account.name or '').lower()
        acc_type = account.account_type

        # Priority 1: Name-based checks
        if any(x in name for x in ['sale', 'revenue', 'service']):
            return 'Sales Accounts'
        if 'purchase' in name or 'cost of goods' in name or 'cogs' in name:
            return 'Purchase Accounts'
        
        # Check for other direct incomes/expenses by name
        if 'direct income' in name:
            return 'Direct Incomes'
        if 'direct expense' in name:
            return 'Direct Expenses'
            
        if 'indirect income' in name:
            return 'Indirect Incomes'
        if any(x in name for x in ['salary', 'rent', 'admin', 'utilities', 'depreciation']):
            return 'Indirect Expenses'

        # Priority 2: Odoo Type fallback
        if acc_type == 'income':
            return 'Sales Accounts'
        if acc_type == 'income_other':
            return 'Indirect Incomes'
        if acc_type == 'expense_direct_cost':
            return 'Direct Expenses'
        if acc_type in ('expense', 'expense_depreciation'):
            return 'Indirect Expenses'
        
        return 'Miscellaneous Expenses'

    def _get_period_balances(self, date_from, date_to, company_id):
        """
        Calculate P&L from core journal items - NET BALANCES for the period
        Returns natural amounts (positive values)
        """
        balances = defaultdict(float)
        
        pl_account_types = [
            'income', 'income_other', 
            'expense_direct_cost', 'expense', 'expense_depreciation'
        ]
        
        accounts = self.env['account.account'].search([
            ('account_type', 'in', pl_account_types),
            ('company_id', '=', company_id.id)
        ])
        
        if not accounts:
            return balances

        account_type_map = {acc.id: acc.account_type for acc in accounts}
        
        domain = [
            ('move_id.state', '=', 'posted'),
            ('date', '>=', date_from),
            ('date', '<=', date_to),
            ('company_id', '=', company_id.id),
            ('account_id', 'in', accounts.ids)
        ]

        read_group_result = self.env['account.move.line'].read_group(
            domain,
            ['debit', 'credit', 'account_id'],
            ['account_id']
        )
        
        income_types = ('income', 'income_other')

        for res in read_group_result:
            if not res['account_id']:
                continue
                
            account_id = res['account_id'][0]
            account_type = account_type_map.get(account_id)
            
            if not account_type:
                continue

            debit = res['debit'] or 0.0
            credit = res['credit'] or 0.0

            if account_type in income_types:
                # Income: Credit - Debit
                balances[account_id] = credit - debit
            else:
                # Expense: Debit - Credit
                balances[account_id] = debit - credit

        return balances

    def _prepare_report_lines(self):
        """Prepare P&L in Tally standard format"""
        self.ensure_one()
        self.line_ids.unlink()
        
        period_balances = self._get_period_balances(
            self.start_date, self.end_date, self.company_id
        )
        
        if not period_balances:
            self.env['tally.profit.loss.line'].create([{
                'wizard_id': self.id,
                'sequence': 10,
                'level': 0,
                'name': 'No transactions in this period',
                'amount': 0.0,
                'is_group': False,
                'is_net_result': False,
            }])
            return
        
        all_accounts = self.env['account.account'].browse(period_balances.keys())
        
        # Group accounts
        accounts_by_group = defaultdict(lambda: self.env['account.account'])
        
        for account in all_accounts:
            group_name = self._classify_pl_account_to_tally_group(account)
            accounts_by_group[group_name] |= account
        
        expense_groups = [
            'Purchase Accounts',
            'Direct Expenses',
            'Indirect Expenses',
            'Miscellaneous Expenses'
        ]
        
        income_groups = [
            'Sales Accounts',
            'Direct Incomes',
            'Indirect Incomes'
        ]
        
        lines = []
        sequence = 0
        
        total_expenses = 0.0
        total_income = 0.0
        
        # Process Income Groups
        for group_name in income_groups:
            accounts = accounts_by_group.get(group_name)
            if not accounts:
                continue
            
            group_total = 0.0
            group_lines = []
            
            for account in sorted(accounts, key=lambda a: (a.code or '', a.name)):
                # balance is (credit - debit)
                balance = period_balances.get(account.id, 0.0)
                
                if abs(balance) < 0.01:
                    continue
                
                group_lines.append({
                    'level': 1,
                    'name': f"  {account.name}",
                    'code': account.code,
                    'amount': balance, # Keep positive
                    'is_group': False,
                    'is_total': False,
                })
                
                group_total += balance
            
            if group_lines:
                sequence += 10
                lines.append({
                    'wizard_id': self.id,
                    'sequence': sequence,
                    'level': 0,
                    'name': group_name,
                    'amount': group_total,
                    'is_group': True,
                })
                
                for line_vals in group_lines:
                    sequence += 10
                    line_vals.update({
                        'wizard_id': self.id,
                        'sequence': sequence
                    })
                    lines.append(line_vals)
                
                total_income += group_total
        
        # Process Expense Groups
        for group_name in expense_groups:
            accounts = accounts_by_group.get(group_name)
            if not accounts:
                continue
            
            group_total = 0.0
            group_lines = []
            
            for account in sorted(accounts, key=lambda a: (a.code or '', a.name)):
                # balance is (debit - credit)
                balance = period_balances.get(account.id, 0.0)
                
                if abs(balance) < 0.01:
                    continue
                
                group_lines.append({
                    'level': 1,
                    'name': f"  {account.name}",
                    'code': account.code,
                    'amount': balance, # Keep positive
                    'is_group': False,
                    'is_total': False,
                })
                
                group_total += balance
            
            if group_lines:
                sequence += 10
                lines.append({
                    'wizard_id': self.id,
                    'sequence': sequence,
                    'level': 0,
                    'name': group_name,
                    'amount': group_total,
                    'is_group': True,
                })
                
                for line_vals in group_lines:
                    sequence += 10
                    line_vals.update({
                        'wizard_id': self.id,
                        'sequence': sequence
                    })
                    lines.append(line_vals)
                
                total_expenses += group_total

        # Net Profit/Loss
        net_result = total_income - total_expenses
        
        sequence += 10
        if net_result >= 0:
            lines.append({
                'wizard_id': self.id,
                'sequence': sequence,
                'level': 0,
                'name': 'Net Profit',
                'amount': net_result,
                'is_net_result': True,
            })
        else:
            lines.append({
                'wizard_id': self.id,
                'sequence': sequence,
                'level': 0,
                'name': 'Net Loss',
                'amount': abs(net_result),
                'is_net_result': True,
            })

        self.env['tally.profit.loss.line'].create(lines)

    def action_view_report(self):
        self.ensure_one()
        self._prepare_report_lines()
        self.file_name = f"Profit_Loss_{self.company_id.name}_{self.start_date}_to_{self.end_date}"
        return self.env.ref('accounting_excel_reports.action_report_tally_profit_loss').report_action(self)

    def action_download_excel(self):
        self.ensure_one()
        if not self.line_ids:
            self._prepare_report_lines()

        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet('Profit & Loss')

        formats = {
            'title': workbook.add_format({
                'bold': True, 'font_size': 14, 'align': 'center', 'font_name': 'Arial'
            }),
            'subtitle': workbook.add_format({
                'align': 'center', 'font_size': 10, 'font_name': 'Arial'
            }),
            'header': workbook.add_format({
                'bold': True, 'border': 1, 'align': 'center',
                'bg_color': '#D3D3D3', 'font_name': 'Arial'
            }),
            'group': workbook.add_format({
                'bold': True, 'font_name': 'Arial', 'font_size': 10
            }),
            'group_number': workbook.add_format({
                'bold': True, 'num_format': '#,##0.00', 'font_name': 'Arial'
            }),
            'account': workbook.add_format({
                'font_name': 'Arial', 'font_size': 9
            }),
            'number': workbook.add_format({
                'num_format': '#,##0.00', 'font_name': 'Arial', 'font_size': 9
            }),
            'profit': workbook.add_format({
                'bold': True, 'top': 2, 'bottom': 6,
                'num_format': '#,##0.00', 'bg_color': '#C6EFCE',
                'font_color': '#006100', 'font_name': 'Arial'
            }),
            'loss': workbook.add_format({
                'bold': True, 'top': 2, 'bottom': 6,
                'num_format': '#,##0.00', 'bg_color': '#FFC7CE',
                'font_color': '#9C0006', 'font_name': 'Arial'
            }),
        }

        worksheet.merge_range('A1:B1', self.company_id.name, formats['title'])
        worksheet.merge_range('A2:B2', 'Profit & Loss Account', formats['title'])
        worksheet.merge_range('A3:B3', 
            f'From {self.start_date.strftime("%d-%b-%Y")} To {self.end_date.strftime("%d-%b-%Y")}',
            formats['subtitle']
        )

        worksheet.set_column('A:A', 50)
        worksheet.set_column('B:B', 18)

        row = 4
        worksheet.write(row, 0, 'Particulars', formats['header'])
        worksheet.write(row, 1, 'Amount', formats['header'])

        row = 5
        for line in self.line_ids:
            # We now display amounts as they are (Income positive, Expense positive)
            # Net Profit/Loss logic handles the final sign
            amount_to_display = line.amount
            
            if line.is_net_result:
                fmt = formats['profit'] if 'Profit' in line.name else formats['loss']
                worksheet.write(row, 0, line.name, fmt)
                worksheet.write(row, 1, amount_to_display, fmt)
            elif line.is_group:
                worksheet.write(row, 0, line.name, formats['group'])
                worksheet.write(row, 1, amount_to_display if amount_to_display else '', formats['group_number'])
            else:
                worksheet.write(row, 0, line.name, formats['account'])
                worksheet.write(row, 1, amount_to_display if abs(amount_to_display) > 0.01 else '', formats['number'])
            row += 1

        workbook.close()
        output.seek(0)

        excel_data = output.read()
        self.excel_file = base64.b64encode(excel_data)
        self.file_name = f'Profit_Loss_{self.start_date.strftime("%d%m%Y")}_{self.end_date.strftime("%d%m%Y")}.xlsx'

        return {
            'type': 'ir.actions.act_url',
            'url': f'/web/content?model=profit.loss.wizard&id={self.id}&field=excel_file&filename_field=file_name&download=true',
            'target': 'self',
        }