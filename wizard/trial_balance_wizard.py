from odoo import models, fields, api
from odoo.exceptions import UserError
import base64
from io import BytesIO
import xlsxwriter
from collections import defaultdict

class TrialBalanceWizard(models.TransientModel):
    _name = 'trial.balance.wizard'
    _description = 'Trial Balance Report Wizard'

    start_date = fields.Date(string='Start Date', required=True,
                             help="Not used for Tally-style Closing Balance, but good for context.")
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

    def _get_account_balances(self, date_to, company_id):
        """
        Calculate account balances from Odoo's core journal items.
        Returns Debit-Credit format (can be positive or negative):
        - Positive balance = Debit balance
        - Negative balance = Credit balance
        """
        balances = defaultdict(float)
        
        # Domain for all posted journal items up to the end date
        domain = [
            ('move_id.state', '=', 'posted'),
            ('date', '<=', date_to),
            ('company_id', '=', company_id.id),
            ('account_id.account_type', '!=', 'off_balance') # Exclude off-balance accounts
        ]

        # Use read_group to sum debit and credit by account
        read_group_result = self.env['account.move.line'].read_group(
            domain,
            ['debit', 'credit', 'account_id'],
            ['account_id']
        )

        # Calculate the balance (debit - credit) for each account
        for res in read_group_result:
            if res['account_id']:
                account_id = res['account_id'][0]
                balances[account_id] += (res['debit'] or 0.0) - (res['credit'] or 0.0)

        return balances

    def _prepare_report_lines(self):
        """Prepare report lines in Tally standard format"""
        self.ensure_one()
        self.line_ids.unlink()
        lines = []
        sequence = 0

        # Use the new core calculation method
        account_balances = self._get_account_balances(self.end_date, self.company_id)
        
        # We only need to process accounts that have a balance
        if not account_balances:
            return

        all_accounts = self.env['account.account'].browse(account_balances.keys())

        # Tally Standard Grouping - Proper Order
        TALLY_GROUP_MAP = {
            'Capital Account': ['equity', 'equity_unaffected'],
            'Current Liabilities': ['liability_payable', 'liability_credit_card', 'liability_current'],
            'Loans (Liability)': ['liability_non_current'],
            'Fixed Assets': ['asset_fixed', 'asset_non_current'],
            'Investments': [], # Add 'asset_investment' if you use it
            'Current Assets': ['asset_receivable', 'asset_cash', 'asset_current', 'asset_prepayment'],
            'Loans & Advances (Asset)': [], # Placeholder for Tally group
            'Suspense A/c': [], # Placeholder for Tally group
            'Sales Accounts': ['income'],
            'Purchase Accounts': ['expense_direct_cost'],
            'Direct Incomes': [], # Placeholder for Tally group
            'Direct Expenses': ['expense'],
            'Indirect Incomes': ['income_other'],
            'Indirect Expenses': ['expense_depreciation'],
        }

        type_to_group_map = {}
        for group_name, type_list in TALLY_GROUP_MAP.items():
            for acc_type in type_list:
                type_to_group_map[acc_type] = group_name

        accounts_by_group = defaultdict(lambda: self.env['account.account'])
        
        # Assign accounts to groups
        for acc in all_accounts:
            # Use 'Miscellaneous' as a fallback group if no type is mapped
            group_name = type_to_group_map.get(acc.account_type, 'Miscellaneous')
            accounts_by_group[group_name] |= acc

        # Tally Standard Order
        group_order = [
            'Capital Account',
            'Loans (Liability)',
            'Current Liabilities',
            'Fixed Assets',
            'Investments',
            'Current Assets',
            'Loans & Advances (Asset)',
            'Suspense A/c',
            'Sales Accounts',
            'Purchase Accounts',
            'Direct Incomes',
            'Direct Expenses',
            'Indirect Incomes',
            'Indirect Expenses',
            'Miscellaneous' # Add fallback group
        ]

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

                # TALLY STANDARD: Positive = Debit, Negative = Credit
                debit = balance if balance > 0 else 0.0
                credit = abs(balance) if balance < 0 else 0.0
                
                # --- CHANGE ---
                # Don't add sequence here. Just prepare the line data.
                group_lines.append({
                    'level': 1,
                    'name': f"  {account.name} ({account.code or 'N/A'})",
                    'debit': debit,
                    'credit': credit,
                    'is_group': False,
                    'is_total': False,
                })
                group_debit_total += debit
                group_credit_total += credit

            if group_lines:
                # ---
                # CHANGE: Add the group header line FIRST, then the accounts.
                # This ensures a strictly increasing sequence.
                # ---

                # 1. Add the group header line
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
                
                # 2. Add the account lines for this group
                for line_vals in group_lines:
                    sequence += 10 # Increment sequence for each account
                    line_vals.update({
                        'wizard_id': self.id,
                        'sequence': sequence
                    })
                    lines.append(line_vals)

                # 3. Update grand totals
                grand_total_debit += group_debit_total
                grand_total_credit += group_credit_total

        # Grand Total (must always match: Debit = Credit)
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

        # Tally-style formats
        formats = {
            'title': workbook.add_format({
                'bold': True, 
                'font_size': 14, 
                'align': 'center',
                'font_name': 'Arial'
            }),
            'subtitle': workbook.add_format({
                'align': 'center', 
                'font_size': 10,
                'font_name': 'Arial'
            }),
            'header': workbook.add_format({
                'bold': True, 
                'border': 1,
                'align': 'center',
                'bg_color': '#D3D3D3',
                'font_name': 'Arial'
            }),
            'group': workbook.add_format({
                'bold': True,
                'font_name': 'Arial',
                'font_size': 10
            }),
            'group_number': workbook.add_format({
                'bold': True,
                'num_format': '#,##0.00',
                'font_name': 'Arial'
            }),
            'account': workbook.add_format({
                'font_name': 'Arial',
                'font_size': 9
            }),
            'number': workbook.add_format({
                'num_format': '#,##0.00',
                'font_name': 'Arial',
                'font_size': 9
            }),
            'total': workbook.add_format({
                'bold': True,
                'top': 2,
                'bottom': 6,
                'num_format': '#,##0.00',
                'font_name': 'Arial'
            }),
            'total_text': workbook.add_format({
                'bold': True,
                'top': 2,
                'bottom': 6,
                'font_name': 'Arial'
            }),
        }

        # Title
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