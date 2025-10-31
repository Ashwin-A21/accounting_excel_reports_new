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

    def _get_period_balances(self, date_from, date_to, company_id):
        """
        Calculate P&L from core journal items - TALLY STANDARD
        Returns balances as NATURAL AMOUNTS (all positive):
        - Income accounts: Positive (Credit balance)
        - Expense accounts: Positive (Debit balance)
        """
        balances = defaultdict(float)
        
        # Define P&L account types
        pl_account_types = [
            'income', 'income_other', 
            'expense_direct_cost', 'expense', 'expense_depreciation'
        ]
        
        # Find all relevant accounts
        accounts = self.env['account.account'].search([
            ('account_type', 'in', pl_account_types),
            ('company_id', '=', company_id.id)
        ])
        if not accounts:
            return balances

        # Create a map of account_id to its type
        account_type_map = {acc.id: acc.account_type for acc in accounts}
        
        # Domain for all posted journal items *within the period*
        domain = [
            ('move_id.state', '=', 'posted'),
            ('date', '>=', date_from),
            ('date', '<=', date_to),
            ('company_id', '=', company_id.id),
            ('account_id', 'in', accounts.ids)
        ]

        # Use read_group to sum debit and credit by account
        read_group_result = self.env['account.move.line'].read_group(
            domain,
            ['debit', 'credit', 'account_id'],
            ['account_id']
        )
        
        # Income types (natural credit balance)
        income_types = ('income', 'income_other')

        # Calculate natural balances
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
                # Income: Natural balance is Credit - Debit
                balances[account_id] = credit - debit
            else:
                # Expense: Natural balance is Debit - Credit
                balances[account_id] = debit - credit

        return balances

    def _prepare_report_lines(self):
        """Prepare P&L in Tally standard format with CLEAR POSITIVE VALUES"""
        self.ensure_one()
        self.line_ids.unlink()
        lines = []
        sequence = 0

        Account = self.env['account.account']
        period_balances = self._get_period_balances(
            self.start_date, self.end_date, self.company_id
        )

        def _create_lines_for_type(account_types, level):
            """Helper to generate line data - returns POSITIVE amounts"""
            account_lines = []
            group_total = 0.0
            
            accounts = Account.search([
                ('account_type', 'in', account_types),
                ('company_id', '=', self.company_id.id)
            ])
            
            for account in sorted(accounts, key=lambda a: (a.code or '', a.name)):
                balance = period_balances.get(account.id, 0.0)
                if abs(balance) < 0.01:
                    continue
                
                account_lines.append({
                    'level': level,
                    'name': f"{'  ' * level}{account.name} ({account.code or 'N/A'})",
                    'code': account.code,
                    'amount': abs(balance),  # Always positive
                    'is_group': False,
                    'is_total': False,
                })
                group_total += abs(balance)
            
            return account_lines, group_total

        # === EXPENSE SIDE (LEFT/DEBIT SIDE IN TALLY) ===
        
        # Purchase Accounts (COGS)
        sequence += 10
        lines.append({
            'wizard_id': self.id, 'sequence': sequence, 'level': 0,
            'name': 'Purchase Accounts', 'is_group': True, 'amount': 0.0,
        })
        purchase_idx = len(lines) - 1
        
        account_lines, purchase_total = _create_lines_for_type(['expense_direct_cost'], 1)
        for line_vals in account_lines:
            sequence += 10
            line_vals.update({'wizard_id': self.id, 'sequence': sequence})
            lines.append(line_vals)
        lines[purchase_idx]['amount'] = purchase_total

        # Direct Expenses
        sequence += 10
        lines.append({
            'wizard_id': self.id, 'sequence': sequence, 'level': 0,
            'name': 'Direct Expenses', 'is_group': True, 'amount': 0.0,
        })
        direct_exp_idx = len(lines) - 1
        
        account_lines, direct_exp_total = _create_lines_for_type(['expense'], 1)
        for line_vals in account_lines:
            sequence += 10
            line_vals.update({'wizard_id': self.id, 'sequence': sequence})
            lines.append(line_vals)
        lines[direct_exp_idx]['amount'] = direct_exp_total

        # Indirect Expenses
        sequence += 10
        lines.append({
            'wizard_id': self.id, 'sequence': sequence, 'level': 0,
            'name': 'Indirect Expenses', 'is_group': True, 'amount': 0.0,
        })
        indirect_exp_idx = len(lines) - 1
        
        account_lines, indirect_exp_total = _create_lines_for_type(['expense_depreciation'], 1)
        for line_vals in account_lines:
            sequence += 10
            line_vals.update({'wizard_id': self.id, 'sequence': sequence})
            lines.append(line_vals)
        lines[indirect_exp_idx]['amount'] = indirect_exp_total

        # Total Expenses
        total_expenses = purchase_total + direct_exp_total + indirect_exp_total
        
        # === INCOME SIDE (RIGHT/CREDIT SIDE IN TALLY) ===
        
        # Sales Accounts
        sequence += 10
        lines.append({
            'wizard_id': self.id, 'sequence': sequence, 'level': 0,
            'name': 'Sales Accounts', 'is_group': True, 'amount': 0.0,
        })
        sales_idx = len(lines) - 1
        
        account_lines, sales_total = _create_lines_for_type(['income'], 1)
        for line_vals in account_lines:
            sequence += 10
            line_vals.update({'wizard_id': self.id, 'sequence': sequence})
            lines.append(line_vals)
        lines[sales_idx]['amount'] = sales_total

        # Direct Incomes
        sequence += 10
        lines.append({
            'wizard_id': self.id, 'sequence': sequence, 'level': 0,
            'name': 'Direct Incomes', 'is_group': True, 'amount': 0.0,
        })
        # Note: No 'direct_income' type by default, add accounts if needed

        # Indirect Incomes
        sequence += 10
        lines.append({
            'wizard_id': self.id, 'sequence': sequence, 'level': 0,
            'name': 'Indirect Incomes', 'is_group': True, 'amount': 0.0,
        })
        indirect_inc_idx = len(lines) - 1
        
        account_lines, indirect_income_total = _create_lines_for_type(['income_other'], 1)
        for line_vals in account_lines:
            sequence += 10
            line_vals.update({'wizard_id': self.id, 'sequence': sequence})
            lines.append(line_vals)
        lines[indirect_inc_idx]['amount'] = indirect_income_total

        # Total Income
        total_income = sales_total + indirect_income_total

        # === NET PROFIT/LOSS (TALLY STANDARD CALCULATION) ===
        # Both are POSITIVE values now
        # Net Profit = Total Income - Total Expenses
        # If Income > Expenses = Profit (shown on Debit/Expense side to balance)
        # If Expenses > Income = Loss (shown on Credit/Income side to balance)
        
        net_result = total_income - total_expenses
        
        sequence += 10
        if net_result >= 0:  # PROFIT
            # Profit is shown on DEBIT side (Expense side) in Tally
            lines.append({
                'wizard_id': self.id, 'sequence': sequence, 'level': 0,
                'name': 'Net Profit', 'amount': net_result,
                'is_net_result': True,
            })
        else:  # LOSS
            # Loss is shown on CREDIT side (Income side) in Tally
            lines.append({
                'wizard_id': self.id, 'sequence': sequence, 'level': 0,
                'name': 'Net Loss', 'amount': abs(net_result),
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
            if line.is_net_result:
                fmt = formats['profit'] if 'Profit' in line.name else formats['loss']
                worksheet.write(row, 0, line.name, fmt)
                worksheet.write(row, 1, line.amount, fmt)
            elif line.is_group:
                worksheet.write(row, 0, line.name, formats['group'])
                worksheet.write(row, 1, line.amount if line.amount else '', formats['group_number'])
            else:
                worksheet.write(row, 0, line.name, formats['account'])
                worksheet.write(row, 1, line.amount if line.amount > 0.01 else '', formats['number'])
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