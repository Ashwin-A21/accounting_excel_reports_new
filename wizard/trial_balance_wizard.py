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
        Intelligently classify accounts into Tally-style groups
        Based on account type, nature, and reconciliation settings
        """
        account_type = account.account_type
        code = (account.code or '').strip()
        name = (account.name or '').lower()
        
        # Receivables → Sundry Debtors
        if account.account_type == 'asset_receivable' or account.reconcile:
            if 'receivable' in name or 'debtor' in name or 'customer' in name:
                return 'Sundry Debtors'
        
        # Payables → Sundry Creditors
        if account.account_type == 'liability_payable' or account.reconcile:
            if 'payable' in name or 'creditor' in name or 'supplier' in name or 'vendor' in name:
                return 'Sundry Creditors'
        
        # Cash & Bank
        if account_type in ('asset_cash', 'asset_current'):
            if any(x in name for x in ['cash', 'bank', 'petty']):
                return 'Cash-in-Hand' if 'cash' in name or 'petty' in name else 'Bank Accounts'
        
        # Capital Account (Equity)
        if account_type in ('equity', 'equity_unaffected'):
            return 'Capital Account'
        
        # Current Liabilities
        if account_type in ('liability_current', 'liability_credit_card'):
            if 'tax' in name or 'gst' in name or 'vat' in name:
                return 'Duties & Taxes'
            return 'Current Liabilities'
        
        # Loans
        if account_type == 'liability_non_current':
            if 'loan' in name or 'borrowing' in name:
                return 'Loans (Liability)'
            return 'Current Liabilities'
        
        # Fixed Assets
        if account_type in ('asset_fixed', 'asset_non_current'):
            return 'Fixed Assets'
        
        # Current Assets
        if account_type in ('asset_current', 'asset_prepayment'):
            if 'inventory' in name or 'stock' in name:
                return 'Stock-in-Hand'
            if 'deposit' in name or 'advance' in name:
                return 'Deposits (Asset)'
            return 'Current Assets'
        
        # Income accounts
        if account_type == 'income':
            return 'Sales Accounts'
        
        if account_type == 'income_other':
            return 'Indirect Incomes'
        
        # Expense accounts
        if account_type == 'expense_direct_cost':
            return 'Purchase Accounts'
        
        if account_type == 'expense':
            return 'Direct Expenses'
        
        if account_type == 'expense_depreciation':
            return 'Indirect Expenses'
        
        return 'Miscellaneous'

    def _get_account_balances(self, date_to, company_id):
        """
        Calculate NET account balances from journal items
        For reconcilable accounts (Debtors/Creditors), only unreconciled items count
        """
        balances = defaultdict(float)
        
        # Get all accounts
        all_accounts = self.env['account.account'].search([
            ('company_id', '=', company_id.id),
            ('account_type', '!=', 'off_balance')
        ])
        
        for account in all_accounts:
            domain = [
                ('account_id', '=', account.id),
                ('move_id.state', '=', 'posted'),
                ('date', '<=', date_to),
                ('company_id', '=', company_id.id)
            ]
            
            # For reconcilable accounts (Debtors/Creditors), only count unreconciled items
            if account.reconcile:
                domain.append(('reconciled', '=', False))
            
            # Sum using read_group
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

    def _prepare_report_lines(self):
        """Prepare Trial Balance in Tally standard format with intelligent grouping"""
        self.ensure_one()
        self.line_ids.unlink()
        
        account_balances = self._get_account_balances(self.end_date, self.company_id)
        
        if not account_balances:
            return
        
        all_accounts = self.env['account.account'].browse(account_balances.keys())
        
        # Group accounts by Tally classification
        accounts_by_group = defaultdict(lambda: self.env['account.account'])
        
        for account in all_accounts:
            group_name = self._classify_account_to_tally_group(account)
            accounts_by_group[group_name] |= account
        
        # Tally Standard Group Order
        group_order = [
            'Capital Account',
            'Current Liabilities',
            'Loans (Liability)',
            'Sundry Creditors',
            'Duties & Taxes',
            'Fixed Assets',
            'Current Assets',
            'Stock-in-Hand',
            'Sundry Debtors',
            'Cash-in-Hand',
            'Bank Accounts',
            'Deposits (Asset)',
            'Sales Accounts',
            'Purchase Accounts',
            'Direct Expenses',
            'Indirect Expenses',
            'Indirect Incomes',
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
            
            for account in sorted(accounts, key=lambda a: (a.code or '', a.name)):
                balance = account_balances.get(account.id, 0.0)
                
                if abs(balance) < 0.01:
                    continue
                
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
                # Add group header
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
                
                # Add account lines
                for line_vals in group_lines:
                    sequence += 10
                    line_vals.update({
                        'wizard_id': self.id,
                        'sequence': sequence
                    })
                    lines.append(line_vals)
                
                grand_total_debit += group_debit_total
                grand_total_credit += group_credit_total
        
        # Grand Total
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

        worksheet.merge_range('A1:C1', self.company_id.name, formats['title'])
        worksheet.merge_range('A2:C2', 'Trial Balance', formats['title'])
        worksheet.merge_range('A3:C3', f'As on {self.end_date.strftime("%d-%b-%Y")}', formats['subtitle'])

        worksheet.set_column('A:A', 50)
        worksheet.set_column('B:B', 18)
        worksheet.set_column('C:C', 18)

        row = 4
        worksheet.write(row, 0, 'Particulars', formats['header'])
        worksheet.write(row, 1, 'Debit', formats['header'])
        worksheet.write(row, 2, 'Credit', formats['header'])

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