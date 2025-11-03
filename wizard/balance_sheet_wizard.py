from odoo import models, fields, api
from odoo.exceptions import UserError
import xlsxwriter
from io import BytesIO
import base64
from collections import defaultdict

class BalanceSheetWizard(models.TransientModel):
    _name = 'balance.sheet.wizard'
    _description = 'Balance Sheet Report Wizard'
    
    end_date = fields.Date(string='As on Date', required=True)
    company_id = fields.Many2one('res.company', string='Company', 
                                 required=True, 
                                 default=lambda self: self.env.company)
    excel_file = fields.Binary(string='Excel File', readonly=True)
    file_name = fields.Char(string='File Name', readonly=True)
    line_ids = fields.One2many('tally.balance.sheet.line', 'wizard_id', string='Report Lines (Vertical)')
    horizontal_view = fields.Boolean(string="Horizontal View")
    liability_line_ids = fields.One2many('tally.balance.sheet.line', 'wizard_liab_id', string='Liability Lines')
    asset_line_ids = fields.One2many('tally.balance.sheet.line', 'wizard_asset_id', string='Asset Lines')

    def _classify_bs_account_to_tally_group(self, account):
        """Classify Balance Sheet accounts into Tally-style groups"""
        account_type = account.account_type
        name = (account.name or '').lower()
        
        # === LIABILITIES ===
        if account_type in ('equity', 'equity_unaffected'):
            return 'Capital Account'
        
        if account_type == 'liability_payable' or account.reconcile:
            if 'payable' in name or 'creditor' in name or 'supplier' in name:
                return 'Sundry Creditors'
        
        if account_type in ('liability_current', 'liability_credit_card'):
            if 'tax' in name or 'gst' in name or 'vat' in name:
                return 'Duties & Taxes'
            if 'provision' in name:
                return 'Provisions'
            return 'Current Liabilities'
        
        if account_type == 'liability_non_current':
            if 'loan' in name or 'borrowing' in name:
                return 'Loans (Liability)'
            return 'Current Liabilities'
        
        # === ASSETS ===
        if account_type in ('asset_fixed', 'asset_non_current'):
            return 'Fixed Assets'
        
        if account_type == 'asset_receivable' or account.reconcile:
            if 'receivable' in name or 'debtor' in name or 'customer' in name:
                return 'Sundry Debtors'
        
        if account_type == 'asset_cash':
            if 'cash' in name or 'petty' in name:
                return 'Cash-in-Hand'
            return 'Bank Accounts'
        
        if account_type in ('asset_current', 'asset_prepayment'):
            if 'inventory' in name or 'stock' in name:
                return 'Stock-in-Hand'
            if 'deposit' in name or 'advance' in name:
                return 'Deposits (Asset)'
            if 'bank' in name:
                return 'Bank Accounts'
            return 'Current Assets'
        
        return 'Miscellaneous'

    def _get_closing_balances(self, date_to, company_id):
        """
        Get closing balances - NET BALANCES for reconcilable accounts
        """
        balances = defaultdict(float)
        
        bs_account_types = [
            'asset_receivable', 'asset_cash', 'asset_current', 'asset_prepayment',
            'asset_fixed', 'asset_non_current',
            'liability_payable', 'liability_current', 'liability_non_current', 
            'liability_credit_card',
            'equity', 'equity_unaffected'
        ]
        
        accounts = self.env['account.account'].search([
            ('account_type', 'in', bs_account_types),
            ('company_id', '=', company_id.id)
        ])
        
        if not accounts:
            return balances

        for account in accounts:
            domain = [
                ('account_id', '=', account.id),
                ('move_id.state', '=', 'posted'),
                ('date', '<=', date_to),
                ('company_id', '=', company_id.id)
            ]
            
            # For reconcilable accounts, only unreconciled items
            if account.reconcile:
                domain.append(('reconciled', '=', False))
            
            result = self.env['account.move.line'].read_group(
                domain,
                ['debit', 'credit'],
                []
            )
            
            if result:
                debit = result[0].get('debit', 0.0)
                credit = result[0].get('credit', 0.0)
                balance = debit - credit
                
                if abs(balance) >= 0.01:
                    balances[account.id] = balance
        
        return balances

    def _get_period_profit_loss(self, date_from, date_to, company_id):
        """Calculate Net Profit/Loss for the fiscal year"""
        pl_account_types = [
            'income', 'income_other', 
            'expense_direct_cost', 'expense', 'expense_depreciation'
        ]
        
        accounts = self.env['account.account'].search([
            ('account_type', 'in', pl_account_types),
            ('company_id', '=', company_id.id)
        ])
        
        if not accounts:
            return 0.0

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
        
        income_total = 0.0
        expense_total = 0.0
        
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
                income_total += (credit - debit)
            else:
                expense_total += (debit - credit)

        return income_total - expense_total

    def _prepare_vertical_report_lines(self):
        """Prepare Balance Sheet in vertical format"""
        self.ensure_one()
        self.line_ids.unlink()
        
        closing_balances = self._get_closing_balances(self.end_date, self.company_id)
        
        if not closing_balances:
            return
        
        all_accounts = self.env['account.account'].browse(closing_balances.keys())
        
        # Group accounts
        accounts_by_group = defaultdict(lambda: self.env['account.account'])
        
        for account in all_accounts:
            group_name = self._classify_bs_account_to_tally_group(account)
            accounts_by_group[group_name] |= account
        
        # Tally BS Group Order
        liability_groups = [
            'Capital Account',
            'Current Liabilities',
            'Loans (Liability)',
            'Sundry Creditors',
            'Duties & Taxes',
            'Provisions'
        ]
        
        asset_groups = [
            'Fixed Assets',
            'Current Assets',
            'Stock-in-Hand',
            'Sundry Debtors',
            'Cash-in-Hand',
            'Bank Accounts',
            'Deposits (Asset)',
            'Miscellaneous'
        ]
        
        lines = []
        sequence = 0
        
        # Calculate fiscal year P&L
        fiscal_year_start = self.company_id.compute_fiscalyear_dates(self.end_date)['date_from']
        net_profit_loss = self._get_period_profit_loss(
            fiscal_year_start, self.end_date, self.company_id
        )
        
        # === LIABILITIES ===
        total_liabilities = 0.0
        
        for group_name in liability_groups:
            accounts = accounts_by_group.get(group_name)
            if not accounts:
                if group_name == 'Capital Account':
                    # Always show Capital Account even if empty
                    sequence += 10
                    lines.append({
                        'wizard_id': self.id,
                        'sequence': sequence,
                        'level': 0,
                        'name': group_name,
                        'amount': 0.0,
                        'is_group': True,
                    })
                continue
            
            group_total = 0.0
            group_lines = []
            
            for account in sorted(accounts, key=lambda a: (a.code or '', a.name)):
                balance = closing_balances.get(account.id, 0.0)
                
                if abs(balance) < 0.01:
                    continue
                
                # For liabilities, we want credit balance (negative in our system)
                # Convert to positive for display
                amount = abs(balance)
                
                group_lines.append({
                    'level': 1,
                    'name': f"  {account.name}",
                    'code': account.code,
                    'amount': amount,
                    'is_group': False,
                })
                
                # For liability accounts, credit balance is normal
                if balance < 0:
                    group_total += amount
                else:
                    # Debit balance in liability account (unusual)
                    group_total -= amount
            
            if group_lines or group_name == 'Capital Account':
                sequence += 10
                
                # Add P&L to Capital Account
                if group_name == 'Capital Account':
                    group_total += net_profit_loss
                
                lines.append({
                    'wizard_id': self.id,
                    'sequence': sequence,
                    'level': 0,
                    'name': group_name,
                    'amount': abs(group_total),
                    'is_group': True,
                })
                
                for line_vals in group_lines:
                    sequence += 10
                    line_vals.update({
                        'wizard_id': self.id,
                        'sequence': sequence
                    })
                    lines.append(line_vals)
                
                # Add P&L line under Capital
                if group_name == 'Capital Account' and abs(net_profit_loss) > 0.01:
                    sequence += 10
                    pl_name = "  Net Profit" if net_profit_loss >= 0 else "  Net Loss"
                    lines.append({
                        'wizard_id': self.id,
                        'sequence': sequence,
                        'level': 1,
                        'name': pl_name,
                        'amount': abs(net_profit_loss),
                    })
                
                total_liabilities += abs(group_total)
        
        # === ASSETS ===
        total_assets = 0.0
        
        for group_name in asset_groups:
            accounts = accounts_by_group.get(group_name)
            if not accounts:
                continue
            
            group_total = 0.0
            group_lines = []
            
            for account in sorted(accounts, key=lambda a: (a.code or '', a.name)):
                balance = closing_balances.get(account.id, 0.0)
                
                if abs(balance) < 0.01:
                    continue
                
                amount = abs(balance)
                
                group_lines.append({
                    'level': 1,
                    'name': f"  {account.name}",
                    'code': account.code,
                    'amount': amount,
                    'is_group': False,
                })
                
                # For asset accounts, debit balance is normal
                if balance > 0:
                    group_total += amount
                else:
                    # Credit balance in asset account (unusual)
                    group_total -= amount
            
            if group_lines:
                sequence += 10
                lines.append({
                    'wizard_id': self.id,
                    'sequence': sequence,
                    'level': 0,
                    'name': group_name,
                    'amount': abs(group_total),
                    'is_group': True,
                })
                
                for line_vals in group_lines:
                    sequence += 10
                    line_vals.update({
                        'wizard_id': self.id,
                        'sequence': sequence
                    })
                    lines.append(line_vals)
                
                total_assets += abs(group_total)
        
        # Total
        sequence += 10
        balance_total = max(total_liabilities, total_assets)
        lines.append({
            'wizard_id': self.id,
            'sequence': sequence,
            'level': 0,
            'name': 'Total',
            'amount': balance_total,
            'is_total': True,
        })

        self.env['tally.balance.sheet.line'].create(lines)

    def _prepare_horizontal_report_lines(self):
        """Prepare Balance Sheet in horizontal format"""
        self.ensure_one()
        self.liability_line_ids.unlink()
        self.asset_line_ids.unlink()
        
        closing_balances = self._get_closing_balances(self.end_date, self.company_id)
        
        if not closing_balances:
            return
        
        all_accounts = self.env['account.account'].browse(closing_balances.keys())
        
        accounts_by_group = defaultdict(lambda: self.env['account.account'])
        
        for account in all_accounts:
            group_name = self._classify_bs_account_to_tally_group(account)
            accounts_by_group[group_name] |= account
        
        liability_groups = [
            'Capital Account',
            'Current Liabilities',
            'Loans (Liability)',
            'Sundry Creditors',
            'Duties & Taxes',
            'Provisions'
        ]
        
        asset_groups = [
            'Fixed Assets',
            'Current Assets',
            'Stock-in-Hand',
            'Sundry Debtors',
            'Cash-in-Hand',
            'Bank Accounts',
            'Deposits (Asset)',
            'Miscellaneous'
        ]
        
        liab_lines = []
        asset_lines = []
        liab_seq = 0
        asset_seq = 0
        
        fiscal_year_start = self.company_id.compute_fiscalyear_dates(self.end_date)['date_from']
        net_profit_loss = self._get_period_profit_loss(
            fiscal_year_start, self.end_date, self.company_id
        )
        
        total_liabilities = 0.0
        
        # Process Liabilities
        for group_name in liability_groups:
            accounts = accounts_by_group.get(group_name)
            if not accounts:
                if group_name == 'Capital Account':
                    liab_seq += 10
                    liab_lines.append({
                        'wizard_liab_id': self.id,
                        'sequence': liab_seq,
                        'level': 0,
                        'name': group_name,
                        'amount': abs(net_profit_loss),
                        'is_group': True,
                    })
                    
                    if abs(net_profit_loss) > 0.01:
                        liab_seq += 10
                        pl_name = "  Net Profit" if net_profit_loss >= 0 else "  Net Loss"
                        liab_lines.append({
                            'wizard_liab_id': self.id,
                            'sequence': liab_seq,
                            'level': 1,
                            'name': pl_name,
                            'amount': abs(net_profit_loss),
                        })
                    
                    total_liabilities += abs(net_profit_loss)
                continue
            
            group_total = 0.0
            group_lines = []
            
            for account in sorted(accounts, key=lambda a: (a.code or '', a.name)):
                balance = closing_balances.get(account.id, 0.0)
                
                if abs(balance) < 0.01:
                    continue
                
                amount = abs(balance)
                
                group_lines.append({
                    'level': 1,
                    'name': f"  {account.name}",
                    'code': account.code,
                    'amount': amount,
                    'is_group': False,
                })
                
                if balance < 0:
                    group_total += amount
                else:
                    group_total -= amount
            
            if group_lines or group_name == 'Capital Account':
                liab_seq += 10
                
                if group_name == 'Capital Account':
                    group_total += net_profit_loss
                
                liab_lines.append({
                    'wizard_liab_id': self.id,
                    'sequence': liab_seq,
                    'level': 0,
                    'name': group_name,
                    'amount': abs(group_total),
                    'is_group': True,
                })
                
                for line_vals in group_lines:
                    liab_seq += 10
                    line_vals.update({
                        'wizard_liab_id': self.id,
                        'sequence': liab_seq
                    })
                    liab_lines.append(line_vals)
                
                if group_name == 'Capital Account' and abs(net_profit_loss) > 0.01:
                    liab_seq += 10
                    pl_name = "  Net Profit" if net_profit_loss >= 0 else "  Net Loss"
                    liab_lines.append({
                        'wizard_liab_id': self.id,
                        'sequence': liab_seq,
                        'level': 1,
                        'name': pl_name,
                        'amount': abs(net_profit_loss),
                    })
                
                total_liabilities += abs(group_total)
        
        # Process Assets
        total_assets = 0.0
        
        for group_name in asset_groups:
            accounts = accounts_by_group.get(group_name)
            if not accounts:
                continue
            
            group_total = 0.0
            group_lines = []
            
            for account in sorted(accounts, key=lambda a: (a.code or '', a.name)):
                balance = closing_balances.get(account.id, 0.0)
                
                if abs(balance) < 0.01:
                    continue
                
                amount = abs(balance)
                
                group_lines.append({
                    'level': 1,
                    'name': f"  {account.name}",
                    'code': account.code,
                    'amount': amount,
                    'is_group': False,
                })
                
                if balance > 0:
                    group_total += amount
                else:
                    group_total -= amount
            
            if group_lines:
                asset_seq += 10
                asset_lines.append({
                    'wizard_asset_id': self.id,
                    'sequence': asset_seq,
                    'level': 0,
                    'name': group_name,
                    'amount': abs(group_total),
                    'is_group': True,
                })
                
                for line_vals in group_lines:
                    asset_seq += 10
                    line_vals.update({
                        'wizard_asset_id': self.id,
                        'sequence': asset_seq
                    })
                    asset_lines.append(line_vals)
                
                total_assets += abs(group_total)
        
        # Totals
        liab_seq += 10
        liab_lines.append({
            'wizard_liab_id': self.id,
            'sequence': liab_seq,
            'level': 0,
            'name': 'Total',
            'amount': total_liabilities,
            'is_total': True
        })
        
        asset_seq += 10
        asset_lines.append({
            'wizard_asset_id': self.id,
            'sequence': asset_seq,
            'level': 0,
            'name': 'Total',
            'amount': total_assets,
            'is_total': True
        })
        
        if liab_lines:
            self.env['tally.balance.sheet.line'].create(liab_lines)
        if asset_lines:
            self.env['tally.balance.sheet.line'].create(asset_lines)

    def action_view_report(self):
        self.ensure_one()
        self.file_name = f"Balance_Sheet_{self.company_id.name}_as_on_{self.end_date}"

        if self.horizontal_view:
            self._prepare_horizontal_report_lines()
            return self.env.ref('accounting_excel_reports.action_report_tally_balance_sheet_horizontal').report_action(self)
        else:
            self._prepare_vertical_report_lines()
            return self.env.ref('accounting_excel_reports.action_report_tally_balance_sheet').report_action(self)

    def action_download_excel(self):
        self.ensure_one()
        
        if self.horizontal_view:
            if not self.liability_line_ids or not self.asset_line_ids:
                self._prepare_horizontal_report_lines()
            return self._download_horizontal_excel()
        else:
            if not self.line_ids:
                self._prepare_vertical_report_lines()
            return self._download_vertical_excel()

    def _download_vertical_excel(self):
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet('Balance Sheet')

        formats = {
            'title': workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center', 'font_name': 'Arial'}),
            'subtitle': workbook.add_format({'align': 'center', 'font_size': 10, 'font_name': 'Arial'}),
            'header': workbook.add_format({'bold': True, 'border': 1, 'align': 'center', 'bg_color': '#D3D3D3', 'font_name': 'Arial'}),
            'group': workbook.add_format({'bold': True, 'font_name': 'Arial', 'font_size': 10}),
            'group_number': workbook.add_format({'bold': True, 'num_format': '#,##0.00', 'font_name': 'Arial'}),
            'account': workbook.add_format({'font_name': 'Arial', 'font_size': 9}),
            'number': workbook.add_format({'num_format': '#,##0.00', 'font_name': 'Arial', 'font_size': 9}),
            'total': workbook.add_format({'bold': True, 'top': 2, 'bottom': 6, 'num_format': '#,##0.00', 'font_name': 'Arial'}),
            'total_text': workbook.add_format({'bold': True, 'top': 2, 'bottom': 6, 'font_name': 'Arial'}),
        }

        worksheet.merge_range('A1:B1', self.company_id.name, formats['title'])
        worksheet.merge_range('A2:B2', 'Balance Sheet', formats['title'])
        worksheet.merge_range('A3:B3', f'As on {self.end_date.strftime("%d-%b-%Y")}', formats['subtitle'])

        worksheet.set_column('A:A', 50)
        worksheet.set_column('B:B', 18)

        row = 4
        worksheet.write(row, 0, 'Particulars', formats['header'])
        worksheet.write(row, 1, 'Amount', formats['header'])

        row = 5
        for line in self.line_ids:
            if line.is_total:
                worksheet.write(row, 0, line.name, formats['total_text'])
                worksheet.write(row, 1, line.amount, formats['total'])
            elif line.is_group:
                worksheet.write(row, 0, line.name, formats['group'])
                worksheet.write(row, 1, line.amount if line.amount > 0.01 else '', formats['group_number'])
            else:
                worksheet.write(row, 0, line.name, formats['account'])
                worksheet.write(row, 1, line.amount if line.amount > 0.01 else '', formats['number'])
            row += 1

        workbook.close()
        output.seek(0)

        excel_data = output.read()
        self.excel_file = base64.b64encode(excel_data)
        self.file_name = f'Balance_Sheet_{self.end_date.strftime("%d%m%Y")}.xlsx'

        return {
            'type': 'ir.actions.act_url',
            'url': f'/web/content?model=balance.sheet.wizard&id={self.id}&field=excel_file&filename_field=file_name&download=true',
            'target': 'self',
        }

    def _download_horizontal_excel(self):
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet('Balance Sheet')

        formats = {
            'title': workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center', 'font_name': 'Arial'}),
            'subtitle': workbook.add_format({'align': 'center', 'font_size': 10, 'font_name': 'Arial'}),
            'header': workbook.add_format({'bold': True, 'border': 1, 'align': 'center', 'bg_color': '#D3D3D3', 'font_name': 'Arial'}),
            'group': workbook.add_format({'bold': True, 'font_name': 'Arial', 'font_size': 10}),
            'group_number': workbook.add_format({'bold': True, 'num_format': '#,##0.00', 'font_name': 'Arial'}),
            'account': workbook.add_format({'font_name': 'Arial', 'font_size': 9}),
            'number': workbook.add_format({'num_format': '#,##0.00', 'font_name': 'Arial', 'font_size': 9}),
            'total': workbook.add_format({'bold': True, 'top': 2, 'bottom': 6, 'num_format': '#,##0.00', 'font_name': 'Arial'}),
            'total_text': workbook.add_format({'bold': True, 'top': 2, 'bottom': 6, 'font_name': 'Arial'}),
        }

        worksheet.merge_range('A1:D1', self.company_id.name, formats['title'])
        worksheet.merge_range('A2:D2', 'Balance Sheet', formats['title'])
        worksheet.merge_range('A3:D3', f'As on {self.end_date.strftime("%d-%b-%Y")}', formats['subtitle'])

        worksheet.set_column('A:A', 40)
        worksheet.set_column('B:B', 15)
        worksheet.set_column('C:C', 40)
        worksheet.set_column('D:D', 15)

        row = 4
        worksheet.write(row, 0, 'Liabilities', formats['header'])
        worksheet.write(row, 1, 'Amount', formats['header'])
        worksheet.write(row, 2, 'Assets', formats['header'])
        worksheet.write(row, 3, 'Amount', formats['header'])

        row = 5
        max_rows = max(len(self.liability_line_ids), len(self.asset_line_ids))
        
        for i in range(max_rows):
            if i < len(self.liability_line_ids):
                line = self.liability_line_ids[i]
                if line.is_total:
                    worksheet.write(row, 0, line.name, formats['total_text'])
                    worksheet.write(row, 1, line.amount, formats['total'])
                elif line.is_group:
                    worksheet.write(row, 0, line.name, formats['group'])
                    worksheet.write(row, 1, line.amount if line.amount > 0.01 else '', formats['group_number'])
                else:
                    worksheet.write(row, 0, line.name, formats['account'])
                    worksheet.write(row, 1, line.amount if line.amount > 0.01 else '', formats['number'])
            
            if i < len(self.asset_line_ids):
                line = self.asset_line_ids[i]
                if line.is_total:
                    worksheet.write(row, 2, line.name, formats['total_text'])
                    worksheet.write(row, 3, line.amount, formats['total'])
                elif line.is_group:
                    worksheet.write(row, 2, line.name, formats['group'])
                    worksheet.write(row, 3, line.amount if line.amount > 0.01 else '', formats['group_number'])
                else:
                    worksheet.write(row, 2, line.name, formats['account'])
                    worksheet.write(row, 3, line.amount if line.amount > 0.01 else '', formats['number'])
            
            row += 1

        workbook.close()
        output.seek(0)

        excel_data = output.read()
        self.excel_file = base64.b64encode(excel_data)
        self.file_name = f'Balance_Sheet_Horizontal_{self.end_date.strftime("%d%m%Y")}.xlsx'

        return {
            'type': 'ir.actions.act_url',
            'url': f'/web/content?model=balance.sheet.wizard&id={self.id}&field=excel_file&filename_field=file_name&download=true',
            'target': 'self',
        }