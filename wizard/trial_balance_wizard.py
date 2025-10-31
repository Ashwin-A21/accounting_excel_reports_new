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
        """Get closing balances for all accounts up to a specific date."""
        domain = [
            ('date', '<=', date_to),
            ('company_id', '=', company_id.id),
            ('parent_state', '=', 'posted')
        ]

        account_data = self.env['account.move.line'].read_group(
            domain,
            ['account_id', 'debit', 'credit'],
            ['account_id']
        )

        balances = {}
        for data in account_data:
            balance = data['debit'] - data['credit']
            if abs(balance) > 0.001:
                balances[data['account_id'][0]] = balance
        return balances

    def _prepare_report_lines(self):
        """Prepare report lines in Tally standard format"""
        self.ensure_one()
        self.line_ids.unlink()
        lines = []
        sequence = 0

        account_balances = self._get_account_balances(self.end_date, self.company_id)
        all_accounts = self.env['account.account'].browse(account_balances.keys())

        # Tally Standard Grouping - Proper Order
        TALLY_GROUP_MAP = {
            'Capital Account': ['equity', 'equity_unaffected', 'capital'],
            'Current Liabilities': ['liability_payable', 'liability_credit_card', 'liability_current'],
            'Loans (Liability)': ['liability_non_current'],
            'Fixed Assets': ['asset_fixed', 'asset_non_current'],
            'Investments': [],  # If you have investment accounts
            'Current Assets': ['asset_receivable', 'asset_cash', 'asset_current', 'asset_prepayment'],
            'Loans & Advances (Asset)': [],  # If you have loan receivable accounts
            'Suspense A/c': [],  # If you have suspense accounts
            'Sales Accounts': ['income'],
            'Purchase Accounts': ['expense_direct_cost'],
            'Direct Incomes': [],  # If you have direct income accounts
            'Direct Expenses': ['expense'],
            'Indirect Incomes': ['income_other'],
            'Indirect Expenses': ['expense_depreciation'],
        }

        type_to_group_map = {}
        for group_name, type_list in TALLY_GROUP_MAP.items():
            for acc_type in type_list:
                type_to_group_map[acc_type] = group_name

        accounts_by_group = defaultdict(lambda: self.env['account.account'])
        for acc in all_accounts:
            group_name = type_to_group_map.get(acc.account_type, 'Sundry Debtors' if acc.account_type.startswith('asset') else 'Sundry Creditors')
            accounts_by_group[group_name] |= acc

        # Tally Standard Order
        group_order = [
            'Capital Account',
            'Current Liabilities',
            'Loans (Liability)',
            'Sundry Creditors',
            'Fixed Assets',
            'Investments',
            'Current Assets',
            'Loans & Advances (Asset)',
            'Sundry Debtors',
            'Suspense A/c',
            'Sales Accounts',
            'Purchase Accounts',
            'Direct Incomes',
            'Direct Expenses',
            'Indirect Incomes',
            'Indirect Expenses',
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
                if abs(balance) < 0.001:
                    continue

                debit = balance if balance > 0 else 0.0
                credit = -balance if balance < 0 else 0.0
                
                sequence += 10
                group_lines.append({
                    'wizard_id': self.id,
                    'sequence': sequence,
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
                sequence += 10
                lines.append({
                    'wizard_id': self.id,
                    'sequence': sequence - (len(group_lines) * 10) - 5,
                    'level': 0,
                    'name': group_name,
                    'debit': group_debit_total,
                    'credit': group_credit_total,
                    'is_group': True,
                    'is_total': False,
                })
                lines.extend(group_lines)
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