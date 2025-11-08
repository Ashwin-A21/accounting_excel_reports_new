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
        """
        Standalone Tally Classification for Balance Sheet
        This function IGNORES the Odoo COA type and uses Tally logic.
        """
        name = (account.name or '').lower()
        acc_type = account.account_type # Use as a final fallback

        # === LIABILITIES (Priority 1: Name Check) ===
        if 'capital' in name:
            return 'Capital Account'
            
        # --- FIX: 'outstanding payment' is a Current Liability, not a Sundry Creditor ---
        if 'outstanding payment' in name:
            return 'Current Liabilities'
        # --- END FIX ---

        # 'outstanding payment' was removed from this list
        if any(x in name for x in ['creditor', 'payable', 'supplier', 'vendor']):
            return 'Sundry Creditors'
        if any(x in name for x in ['tax', 'gst', 'vat', 'tds']):
            return 'Duties & Taxes'
        if any(x in name for x in ['loan', 'borrowing']):
            return 'Loans (Liability)'
        if 'provision' in name:
            return 'Provisions'

        # === ASSETS (Priority 1: Name Check) ===
        
        # --- FIX: Be explicit about 'outstanding receipt' ---
        if 'outstanding receipt' in name:
            return 'Current Assets'
        # --- END FIX ---
        
        if any(x in name for x in ['debtor', 'receivable', 'customer']):
            return 'Sundry Debtors'
        if 'bank' in name:
            return 'Bank Accounts'
        if 'cash' in name or 'petty' in name:
            return 'Cash-in-Hand'
        if any(x in name for x in ['fixed asset', 'building', 'vehicle', 'machinery', 'furniture']):
            return 'Fixed Assets'
        if any(x in name for x in ['inventory', 'stock']):
            return 'Stock-in-Hand'
        if any(x in name for x in ['deposit', 'prepaid', 'prepayment', 'advance paid']):
            return 'Deposits (Asset)'

        # === Fallback (Priority 2: Odoo Type) ===
        # This only runs if the name checks above fail
        if acc_type in ('equity', 'equity_unaffected'):
            return 'Capital Account'
        if acc_type == 'liability_payable':
            return 'Sundry Creditors'
        if acc_type in ('liability_current', 'liability_credit_card'):
            return 'Current Liabilities' # 'outstanding payment' will also fall back here
        if acc_type == 'liability_non_current':
            return 'Loans (Liability)'
        if acc_type == 'asset_receivable':
            return 'Sundry Debtors'
        if acc_type in ('asset_fixed', 'asset_non_current'):
            return 'Fixed Assets'
        if acc_type == 'asset_cash':
            return 'Bank Accounts'
        if acc_type in ('asset_current', 'asset_prepayment'):
            return 'Current Assets' # 'outstanding receipt' will also fall back here
        
        # Final Fallback
        if acc_type.startswith('liability_'):
            return 'Current Liabilities'
        if acc_type.startswith('asset_'):
            return 'Current Assets'
            
        return 'Miscellaneous'

    def _get_closing_balances(self, date_to, company_id):
        """
        Get closing balances - NET BALANCES (Debit - Credit) for all accounts.
        This is the new logic that IGNORES 'reconciled' status, fixing
        the "payment made but still showing" bug.
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

        # Use read_group for performance
        domain = [
            ('account_id', 'in', accounts.ids),
            ('move_id.state', '=', 'posted'),
            ('date', '<=', date_to),
            ('company_id', '=', company_id.id)
        ]
        
        result_data = self.env['account.move.line'].read_group(
            domain,
            ['debit', 'credit', 'account_id'],
            ['account_id']
        )
        
        for res in result_data:
            if not res['account_id']:
                continue
                
            debit = res.get('debit', 0.0)
            credit = res.get('credit', 0.0)
            balance = debit - credit # This is the true net balance
            
            if abs(balance) >= 0.01:
                balances[res['account_id'][0]] = balance
        
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
                # Income is credit-positive
                income_total += (credit - debit)
            else:
                # Expense is debit-positive
                expense_total += (debit - credit)

        # Profit is Income - Expense
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
            'Loans (Liability)',
            'Current Liabilities',
            'Sundry Creditors',
            'Duties & Taxes',
            'Provisions'
        ]
        
        asset_groups = [
            'Fixed Assets',
            'Current Assets',
            'Stock-in-Hand',
            'Deposits (Asset)',
            'Sundry Debtors',
            'Cash-in-Hand',
            'Bank Accounts',
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
            
            group_total = 0.0
            group_lines = []
            
            # Special case: Always show Capital Account
            if not accounts and group_name != 'Capital Account':
                continue
                
            if accounts:
                for account in sorted(accounts, key=lambda a: (a.code or '', a.name)):
                    # balance = Debit - Credit
                    balance = closing_balances.get(account.id, 0.0)
                    
                    if abs(balance) < 0.01:
                        continue
                    
                    # For liabilities, display Credit balance (balance < 0) as positive
                    amount = balance * -1
                    
                    group_lines.append({
                        'level': 1,
                        'name': f"  {account.name}",
                        'code': account.code,
                        'amount': amount,
                        'is_group': False,
                    })
                    
                    group_total += amount
            
            if group_lines or group_name == 'Capital Account':
                sequence += 10
                
                # Add P&L to Capital Account
                if group_name == 'Capital Account':
                    # net_profit_loss is (Income - Expense)
                    # A profit (positive) increases Liability/Capital
                    # A loss (negative) decreases Liability/Capital
                    # Our 'group_total' is positive, so add profit or subtract loss
                    group_total += net_profit_loss
                
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
                
                total_liabilities += group_total
        
        # === ASSETS ===
        total_assets = 0.0
        
        for group_name in asset_groups:
            accounts = accounts_by_group.get(group_name)
            if not accounts:
                continue
            
            group_total = 0.0
            group_lines = []
            
            for account in sorted(accounts, key=lambda a: (a.code or '', a.name)):
                # balance = Debit - Credit
                balance = closing_balances.get(account.id, 0.0)
                
                if abs(balance) < 0.01:
                    continue
                
                # For assets, display Debit balance (balance > 0) as positive
                amount = balance
                
                group_lines.append({
                    'level': 1,
                    'name': f"  {account.name}",
                    'code': account.code,
                    'amount': amount,
                    'is_group': False,
                })
                
                group_total += amount
            
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
                
                total_assets += group_total
        
        # Total
        sequence += 10
        lines.append({
            'wizard_id': self.id,
            'sequence': sequence,
            'level': 0,
            'name': 'Total (Liabilities)',
            'amount': total_liabilities,
            'is_total': True,
        })
        
        sequence += 10
        lines.append({
            'wizard_id': self.id,
            'sequence': sequence,
            'level': 0,
            'name': 'Total (Assets)',
            'amount': total_assets,
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
            'Loans (Liability)',
            'Current Liabilities',
            'Sundry Creditors',
            'Duties & Taxes',
            'Provisions'
        ]
        
        asset_groups = [
            'Fixed Assets',
            'Current Assets',
            'Stock-in-Hand',
            'Deposits (Asset)',
            'Sundry Debtors',
            'Cash-in-Hand',
            'Bank Accounts',
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
            
            group_total = 0.0
            group_lines = []

            # Special case: Always show Capital Account
            if not accounts and group_name != 'Capital Account':
                continue
            
            if accounts:
                for account in sorted(accounts, key=lambda a: (a.code or '', a.name)):
                    # balance = Debit - Credit
                    balance = closing_balances.get(account.id, 0.0)
                    
                    if abs(balance) < 0.01:
                        continue
                    
                    # Liabilities are (Credit > Debit), so balance < 0
                    # Display as positive
                    amount = balance * -1 
                    
                    group_lines.append({
                        'level': 1,
                        'name': f"  {account.name}",
                        'code': account.code,
                        'amount': amount,
                        'is_group': False,
                    })
                    
                    group_total += amount
            
            if group_lines or group_name == 'Capital Account':
                liab_seq += 10
                
                if group_name == 'Capital Account':
                    group_total += net_profit_loss
                
                liab_lines.append({
                    'wizard_liab_id': self.id,
                    'sequence': liab_seq,
                    'level': 0,
                    'name': group_name,
                    'amount': group_total,
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
                
                total_liabilities += group_total
        
        # Process Assets
        total_assets = 0.0
        
        for group_name in asset_groups:
            accounts = accounts_by_group.get(group_name)
            if not accounts:
                continue
            
            group_total = 0.0
            group_lines = []
            
            for account in sorted(accounts, key=lambda a: (a.code or '', a.name)):
                # balance = Debit - Credit
                balance = closing_balances.get(account.id, 0.0)
                
                if abs(balance) < 0.01:
                    continue
                
                # Assets are (Debit > Credit), so balance > 0
                amount = balance
                
                group_lines.append({
                    'level': 1,
                    'name': f"  {account.name}",
                    'code': account.code,
                    'amount': amount,
                    'is_group': False,
                })
                
                group_total += amount
            
            if group_lines:
                asset_seq += 10
                asset_lines.append({
                    'wizard_asset_id': self.id,
                    'sequence': asset_seq,
                    'level': 0,
                    'name': group_name,
                    'amount': group_total,
                    'is_group': True,
                })
                
                for line_vals in group_lines:
                    asset_seq += 10
                    line_vals.update({
                        'wizard_asset_id': self.id,
                        'sequence': asset_seq
                    })
                    asset_lines.append(line_vals)
                
                total_assets += group_total
        
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
        
        # This is the fix for the unbalanced Total row
        # Pad the shorter list so the totals align in the view
        diff = len(liab_lines) - len(asset_lines)
        if diff > 0:
            for i in range(diff):
                asset_seq += 10
                asset_lines.insert(-1, { # Insert before the total
                    'wizard_asset_id': self.id, 'sequence': asset_seq, 'name': '', 'level': 2
                })
        elif diff < 0:
            for i in range(abs(diff)):
                liab_seq += 10
                liab_lines.insert(-1, { # Insert before the total
                    'wizard_liab_id': self.id, 'sequence': liab_seq, 'name': '', 'level': 2
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
                worksheet.write(row, 1, line.amount if abs(line.amount) > 0.01 else '', formats['group_number'])
            else:
                worksheet.write(row, 0, line.name, formats['account'])
                worksheet.write(row, 1, line.amount if abs(line.amount) > 0.01 else '', formats['number'])
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
        
        liab_lines = self.liability_line_ids
        asset_lines = self.asset_line_ids
        max_rows = max(len(liab_lines), len(asset_lines))
        
        for i in range(max_rows):
            wrote_liab = False
            if i < len(liab_lines):
                line = liab_lines[i]
                if line.name or line.is_total:
                    if line.is_total:
                        worksheet.write(row, 0, line.name, formats['total_text'])
                        worksheet.write(row, 1, line.amount, formats['total'])
                    elif line.is_group:
                        worksheet.write(row, 0, line.name, formats['group'])
                        worksheet.write(row, 1, line.amount if abs(line.amount) > 0.01 else '', formats['group_number'])
                    else:
                        worksheet.write(row, 0, line.name, formats['account'])
                        worksheet.write(row, 1, line.amount if abs(line.amount) > 0.01 else '', formats['number'])
                    wrote_liab = True
            
            wrote_asset = False
            if i < len(asset_lines):
                line = asset_lines[i]
                if line.name or line.is_total:
                    if line.is_total:
                        worksheet.write(row, 2, line.name, formats['total_text'])
                        worksheet.write(row, 3, line.amount, formats['total'])
                    elif line.is_group:
                        worksheet.write(row, 2, line.name, formats['group'])
                        worksheet.write(row, 3, line.amount if abs(line.amount) > 0.01 else '', formats['group_number'])
                    else:
                        worksheet.write(row, 2, line.name, formats['account'])
                        worksheet.write(row, 3, line.amount if abs(line.amount) > 0.01 else '', formats['number'])
                    wrote_asset = True

            if wrote_liab or wrote_asset:
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