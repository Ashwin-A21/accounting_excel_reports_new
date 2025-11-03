from odoo import models, fields, api
from odoo.exceptions import UserError
import base64
from io import BytesIO
import xlsxwriter
from collections import defaultdict

class TrialBalanceWizard(models.TransientModel):
    _name = 'trial.balance.wizard'
    _description = 'Trial Balance Report Wizard'

    start_date = fields.Date(string='Start Date', required=True)
    end_date = fields.Date(string='As on Date', required=True)
    company_id = fields.Many2one('res.company', string='Company',
                                 required=True,
                                 default=lambda self: self.env.company)
    excel_file = fields.Binary(string='Excel File', readonly=True)
    file_name = fields.Char(string='File Name', readonly=True)
    line_ids = fields.One2many('tally.trial.balance.line', 'wizard_id', string='Report Lines')

    @api.constrains('start_date', 'end_date')
    def _check_dates(self):
        for record in self:
            if record.start_date > record.end_date:
                raise UserError('End Date must be greater than Start Date!')

    def _classify_account_to_tally_group(self, account):
        """
        Classify accounts into Tally-style groups based on account type and name
        This matches Tally's standard grouping logic
        """
        account_type = account.account_type
        code = (account.code or '').strip()
        name = (account.name or '').lower()
        
        # === RECEIVABLES (Debtors) ===
        if account.account_type == 'asset_receivable':
            return 'Sundry Debtors'
        
        # === PAYABLES (Creditors) ===
        if account.account_type == 'liability_payable':
            return 'Sundry Creditors'
        
        # === CASH & BANK ===
        if account_type == 'asset_cash':
            if 'cash' in name or 'petty' in name:
                return 'Cash-in-Hand'
            if 'bank' in name:
                return 'Bank Accounts'
            return 'Cash-in-Hand'
        
        # === EQUITY (Capital) ===
        if account_type in ('equity', 'equity_unaffected'):
            return 'Capital Account'
        
        # === CURRENT LIABILITIES ===
        if account_type in ('liability_current', 'liability_credit_card'):
            if any(x in name for x in ['tax', 'gst', 'vat', 'tds', 'duty']):
                return 'Duties & Taxes'
            if 'provision' in name:
                return 'Provisions'
            return 'Current Liabilities'
        
        # === LOANS ===
        if account_type == 'liability_non_current':
            if any(x in name for x in ['loan', 'borrowing', 'debt']):
                return 'Loans (Liability)'
            return 'Current Liabilities'
        
        # === FIXED ASSETS ===
        if account_type in ('asset_fixed', 'asset_non_current'):
            return 'Fixed Assets'
        
        # === CURRENT ASSETS ===
        if account_type in ('asset_current', 'asset_prepayment'):
            if any(x in name for x in ['inventory', 'stock']):
                return 'Stock-in-Hand'
            if any(x in name for x in ['deposit', 'advance', 'prepaid']):
                return 'Deposits (Asset)'
            if 'bank' in name:
                return 'Bank Accounts'
            return 'Current Assets'
        
        # === INCOME ===
        if account_type == 'income':
            if any(x in name for x in ['sale', 'revenue', 'service income']):
                return 'Sales Accounts'
            return 'Direct Incomes'
        
        if account_type == 'income_other':
            return 'Indirect Incomes'
        
        # === EXPENSES ===
        if account_type == 'expense_direct_cost':
            if any(x in name for x in ['purchase', 'cost of goods', 'cogs']):
                return 'Purchase Accounts'
            return 'Direct Expenses'
        
        if account_type == 'expense':
            return 'Direct Expenses'
        
        if account_type == 'expense_depreciation':
            return 'Indirect Expenses'
        
        return 'Miscellaneous'

    def _get_account_balances(self, date_to, company_id):
        """
        Calculate NET account balances from journal items
        IMPROVED LOGIC - Check actual outstanding via amount_residual:
        - Receivables/Payables: Use amount_residual (shows actual unpaid amounts)
        - All other accounts: ALL posted transactions (complete balances)
        
        This is more accurate than checking reconcile flag
        """
        balances = defaultdict(float)
        
        # Get all accounts except off-balance
        all_accounts = self.env['account.account'].search([
            ('company_id', '=', company_id.id),
            ('account_type', '!=', 'off_balance')
        ])
        
        # Separate receivables/payables from other accounts
        rec_pay_accounts = all_accounts.filtered(lambda a: a.account_type in ['asset_receivable', 'liability_payable'])
        other_accounts = all_accounts - rec_pay_accounts
        
        # === HANDLE RECEIVABLES/PAYABLES - Check Outstanding via amount_residual ===
        for account in rec_pay_accounts:
            # Get invoice/bill lines that have amount_residual tracking
            move_lines = self.env['account.move.line'].search([
                ('account_id', '=', account.id),
                ('move_id.state', '=', 'posted'),
                ('date', '<=', date_to),
                ('company_id', '=', company_id.id),
                ('move_id.move_type', 'in', ['out_invoice', 'in_invoice', 'out_refund', 'in_refund', 'out_receipt', 'in_receipt']),
            ])
            
            outstanding_total = 0.0
            for line in move_lines:
                # amount_residual shows unpaid amount
                # Positive for receivables, negative for payables
                if hasattr(line, 'amount_residual_currency'):
                    residual = line.amount_residual_currency or line.amount_residual
                else:
                    residual = line.amount_residual
                
                if abs(residual) > 0.01:
                    outstanding_total += residual
            
            if abs(outstanding_total) > 0.01:
                balances[account.id] = outstanding_total
        
        # === HANDLE OTHER ACCOUNTS - Complete Balances ===
        for account in other_accounts:
            domain = [
                ('account_id', '=', account.id),
                ('move_id.state', '=', 'posted'),
                ('date', '<=', date_to),
                ('company_id', '=', company_id.id)
            ]
            
            # Use read_group for efficient aggregation
            result = self.env['account.move.line'].read_group(
                domain,
                ['debit', 'credit'],
                []
            )
            
            if result:
                debit = result[0].get('debit', 0.0)
                credit = result[0].get('credit', 0.0)
                balance = debit - credit
                
                # Only include accounts with material balance
                if abs(balance) >= 0.01:
                    balances[account.id] = balance
        
        return balances

    def _prepare_report_lines(self):
        """
        Prepare Trial Balance in Tally standard format with intelligent grouping
        Follows Tally's presentation: Groups with totals, then individual accounts
        """
        self.ensure_one()
        self.line_ids.unlink()
        
        account_balances = self._get_account_balances(self.end_date, self.company_id)
        
        if not account_balances:
            # No transactions - create empty report
            self.env['tally.trial.balance.line'].create([{
                'wizard_id': self.id,
                'sequence': 10,
                'level': 0,
                'name': 'No transactions in this period',
                'debit': 0.0,
                'credit': 0.0,
                'is_group': False,
                'is_total': False,
            }])
            return
        
        all_accounts = self.env['account.account'].browse(account_balances.keys())
        
        # Group accounts by Tally classification
        accounts_by_group = defaultdict(lambda: self.env['account.account'])
        
        for account in all_accounts:
            group_name = self._classify_account_to_tally_group(account)
            accounts_by_group[group_name] |= account
        
        # Tally Standard Group Order (Liabilities, Assets, Income, Expenses)
        group_order = [
            # Liabilities
            'Capital Account',
            'Current Liabilities',
            'Loans (Liability)',
            'Sundry Creditors',
            'Duties & Taxes',
            'Provisions',
            # Assets
            'Fixed Assets',
            'Current Assets',
            'Stock-in-Hand',
            'Sundry Debtors',
            'Cash-in-Hand',
            'Bank Accounts',
            'Deposits (Asset)',
            # Income
            'Sales Accounts',
            'Direct Incomes',
            'Indirect Incomes',
            # Expenses
            'Purchase Accounts',
            'Direct Expenses',
            'Indirect Expenses',
            # Other
            'Miscellaneous'
        ]
        
        lines = []
        sequence = 0
        grand_total_debit = 0.0
        grand_total_credit = 0.0
        
        for group_name in group_order:
            accounts = accounts_by_group.get(group_name)
            if not accounts:
                continue
            
            group_debit_total = 0.0
            group_credit_total = 0.0
            group_lines = []
            
            # Process each account in the group
            for account in sorted(accounts, key=lambda a: (a.code or '', a.name)):
                balance = account_balances.get(account.id, 0.0)
                
                if abs(balance) < 0.01:
                    continue
                
                # Split balance into debit/credit for trial balance display
                debit = balance if balance > 0 else 0.0
                credit = abs(balance) if balance < 0 else 0.0
                
                group_lines.append({
                    'level': 1,
                    'name': f"  {account.name}",
                    'debit': debit,
                    'credit': credit,
                    'is_group': False,
                    'is_total': False,
                })
                
                group_debit_total += debit
                group_credit_total += credit
            
            if group_lines:
                # Add group header with totals
                sequence += 10
                lines.append({
                    'wizard_id': self.id,
                    'sequence': sequence,
                    'level': 0,
                    'name': group_name,
                    'debit': group_debit_total,
                    'credit': group_credit_total,
                    'is_group': True,
                    'is_total': False,
                })
                
                # Add individual account lines under the group
                for line_vals in group_lines:
                    sequence += 10
                    line_vals.update({
                        'wizard_id': self.id,
                        'sequence': sequence
                    })
                    lines.append(line_vals)
                
                grand_total_debit += group_debit_total
                grand_total_credit += group_credit_total
        
        # Grand Total (must always balance in a proper Trial Balance)
        sequence += 10
        lines.append({
            'wizard_id': self.id,
            'sequence': sequence,
            'level': 0,
            'name': 'Total',
            'debit': grand_total_debit,
            'credit': grand_total_credit,
            'is_group': False,
            'is_total': True,
        })
        
        self.env['tally.trial.balance.line'].create(lines)

    def action_view_report(self):
        self.ensure_one()
        self._prepare_report_lines()
        self.file_name = f"Trial_Balance_{self.company_id.name}_as_on_{self.end_date}"
        return self.env.ref('accounting_excel_reports.action_report_tally_trial_balance').report_action(self)

    def action_download_excel(self):
        self.ensure_one()
        if not self.line_ids:
            self._prepare_report_lines()

        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet('Trial Balance')

        # Excel formats matching Tally style
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
            'total': workbook.add_format({
                'bold': True, 'top': 2, 'bottom': 6,
                'num_format': '#,##0.00', 'font_name': 'Arial'
            }),
            'total_text': workbook.add_format({
                'bold': True, 'top': 2, 'bottom': 6, 'font_name': 'Arial'
            }),
        }

        # Header section
        worksheet.merge_range('A1:C1', self.company_id.name, formats['title'])
        worksheet.merge_range('A2:C2', 'Trial Balance', formats['title'])
        worksheet.merge_range('A3:C3', f'As on {self.end_date.strftime("%d-%b-%Y")}', formats['subtitle'])

        # Column widths
        worksheet.set_column('A:A', 50)
        worksheet.set_column('B:B', 18)
        worksheet.set_column('C:C', 18)

        # Column headers
        row = 4
        worksheet.write(row, 0, 'Particulars', formats['header'])
        worksheet.write(row, 1, 'Debit', formats['header'])
        worksheet.write(row, 2, 'Credit', formats['header'])

        # Data rows
        row = 5
        for line in self.line_ids:
            if line.is_total:
                worksheet.write(row, 0, line.name, formats['total_text'])
                worksheet.write(row, 1, line.debit, formats['total'])
                worksheet.write(row, 2, line.credit, formats['total'])
            elif line.is_group:
                worksheet.write(row, 0, line.name, formats['group'])
                worksheet.write(row, 1, line.debit, formats['group_number'])
                worksheet.write(row, 2, line.credit, formats['group_number'])
            else:
                worksheet.write(row, 0, line.name, formats['account'])
                worksheet.write(row, 1, line.debit if line.debit else '', formats['number'])
                worksheet.write(row, 2, line.credit if line.credit else '', formats['number'])
            row += 1

        workbook.close()
        output.seek(0)

        excel_data = output.read()
        self.excel_file = base64.b64encode(excel_data)
        self.file_name = f'Trial_Balance_{self.end_date.strftime("%d%m%Y")}.xlsx'

        return {
            'type': 'ir.actions.act_url',
            'url': f'/web/content?model=trial.balance.wizard&id={self.id}&field=excel_file&filename_field=file_name&download=true',
            'target': 'self',
        }