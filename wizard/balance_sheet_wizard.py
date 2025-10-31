from odoo import models, fields, api
from odoo.exceptions import UserError
import xlsxwriter
from io import BytesIO
import base64
from collections import defaultdict

class BalanceSheetWizard(models.TransientModel):
    _name = 'balance.sheet.wizard'
    _description = 'Balance Sheet Report Wizard'

    start_date = fields.Date(string='Start Date', required=True)
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

    @api.constrains('start_date', 'end_date')
    def _check_dates(self):
        for record in self:
            if record.start_date > record.end_date:
                raise UserError('End Date must be greater than Start Date!')

    def _get_all_closing_balances(self, date_to, company_id):
        """Get closing balances for all accounts"""
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
        balances = defaultdict(float)
        for data in account_data:
            balances[data['account_id'][0]] = data['debit'] - data['credit']
        return balances

    def _get_period_profit_loss(self, date_from, date_to):
        """Get Net Profit/Loss for the period"""
        pl_account_types = ['income', 'income_other', 'expense_direct_cost', 'expense', 'expense_depreciation']
        
        domain_pl = [
            ('date', '>=', date_from),
            ('date', '<=', date_to),
            ('company_id', '=', self.company_id.id),
            ('parent_state', '=', 'posted'),
            ('account_id.account_type', 'in', pl_account_types)
        ]
        
        grouped_data = self.env['account.move.line'].read_group(
            domain_pl,
            ['account_id', 'debit', 'credit'],
            ['account_id']
        )
        
        balances = defaultdict(float)
        for data in grouped_data:
            balances[data['account_id'][0]] = data['debit'] - data['credit']
        
        accounts = self.env['account.account'].browse(balances.keys())
        
        total_income = 0.0
        total_expense = 0.0
        income_types = ['income', 'income_other']
        
        for acc in accounts:
            if acc.account_type in income_types:
                total_income += balances[acc.id]
            else:
                total_expense += balances[acc.id]
        
        # Net P/L = Income - Expense (both in D-C terms)
        # Income is negative (Credit), Expense is positive (Debit)
        # Result: negative = profit, positive = loss
        return total_income - total_expense

    def _prepare_vertical_report_lines(self):
        """Prepare Balance Sheet in vertical format - CORRECTED"""
        self.ensure_one()
        self.line_ids.unlink()
        lines = []
        sequence = 0
        Account = self.env['account.account']
        
        closing_balances = self._get_all_closing_balances(self.end_date, self.company_id)

        def _create_lines_for_type(account_types, level):
            account_lines = []
            group_total = 0.0
            accounts = Account.search([
                ('account_type', 'in', account_types),
                ('company_id', '=', self.company_id.id)
            ])
            for account in sorted(accounts, key=lambda a: (a.code or '', a.name)):
                balance = closing_balances.get(account.id, 0.0)
                if abs(balance) < 0.01:
                    continue
                account_lines.append({
                    'level': level,
                    'name': f"{'  ' * level}{account.name}",
                    'code': account.code,
                    'amount': balance,  # Keep raw balance
                    'is_group': False,
                    'is_total': False,
                })
                group_total += balance
            return account_lines, group_total

        # === LIABILITIES ===
        # Capital Account
        sequence += 10
        lines.append({
            'wizard_id': self.id, 'sequence': sequence, 'level': 0,
            'name': 'Capital Account', 'is_group': True, 'amount': 0.0,
        })
        capital_group_idx = len(lines) - 1
        
        account_lines, capital_total = _create_lines_for_type(['equity', 'equity_unaffected', 'capital'], 1)
        for line_vals in account_lines:
            sequence += 10
            line_vals.update({'wizard_id': self.id, 'sequence': sequence})
            lines.append(line_vals)
        
        # Add Current Period P&L
        net_profit_loss = self._get_period_profit_loss(self.start_date, self.end_date)
        sequence += 10
        lines.append({
            'wizard_id': self.id, 'sequence': sequence, 'level': 1,
            'name': "  Profit & Loss A/c", 'amount': net_profit_loss,
        })
        
        total_capital = capital_total + net_profit_loss
        lines[capital_group_idx]['amount'] = total_capital

        # Current Liabilities
        sequence += 10
        lines.append({
            'wizard_id': self.id, 'sequence': sequence, 'level': 0,
            'name': 'Current Liabilities', 'is_group': True, 'amount': 0.0,
        })
        cl_group_idx = len(lines) - 1
        
        account_lines, cl_total = _create_lines_for_type(
            ['liability_payable', 'liability_credit_card', 'liability_current'], 1
        )
        for line_vals in account_lines:
            sequence += 10
            line_vals.update({'wizard_id': self.id, 'sequence': sequence})
            lines.append(line_vals)
        lines[cl_group_idx]['amount'] = cl_total

        # Loans (Liability)
        sequence += 10
        lines.append({
            'wizard_id': self.id, 'sequence': sequence, 'level': 0,
            'name': 'Loans (Liability)', 'is_group': True, 'amount': 0.0,
        })
        ncl_group_idx = len(lines) - 1
        
        account_lines, ncl_total = _create_lines_for_type(['liability_non_current'], 1)
        for line_vals in account_lines:
            sequence += 10
            line_vals.update({'wizard_id': self.id, 'sequence': sequence})
            lines.append(line_vals)
        lines[ncl_group_idx]['amount'] = ncl_total

        total_liabilities = total_capital + cl_total + ncl_total

        # === ASSETS ===
        # Fixed Assets
        sequence += 10
        lines.append({
            'wizard_id': self.id, 'sequence': sequence, 'level': 0,
            'name': 'Fixed Assets', 'is_group': True, 'amount': 0.0,
        })
        fa_group_idx = len(lines) - 1
        
        account_lines, fa_total = _create_lines_for_type(['asset_fixed', 'asset_non_current'], 1)
        for line_vals in account_lines:
            sequence += 10
            line_vals.update({'wizard_id': self.id, 'sequence': sequence})
            lines.append(line_vals)
        lines[fa_group_idx]['amount'] = fa_total

        # Current Assets
        sequence += 10
        lines.append({
            'wizard_id': self.id, 'sequence': sequence, 'level': 0,
            'name': 'Current Assets', 'is_group': True, 'amount': 0.0,
        })
        ca_group_idx = len(lines) - 1
        
        account_lines, ca_total = _create_lines_for_type(
            ['asset_receivable', 'asset_cash', 'asset_current', 'asset_prepayment'], 1
        )
        for line_vals in account_lines:
            sequence += 10
            line_vals.update({'wizard_id': self.id, 'sequence': sequence})
            lines.append(line_vals)
        lines[ca_group_idx]['amount'] = ca_total

        total_assets = fa_total + ca_total

        # === TOTAL ===
        sequence += 10
        lines.append({
            'wizard_id': self.id, 'sequence': sequence, 'level': 0,
            'name': 'Total', 'amount': max(abs(total_liabilities), abs(total_assets)),
            'is_total': True,
        })

        self.env['tally.balance.sheet.line'].create(lines)

    def _prepare_horizontal_report_lines(self):
        """Prepare Balance Sheet in horizontal format - CORRECTED"""
        self.ensure_one()
        self.liability_line_ids.unlink()
        self.asset_line_ids.unlink()
        
        liab_lines = []
        asset_lines = []
        liab_seq = 0
        asset_seq = 0
        
        Account = self.env['account.account']
        closing_balances = self._get_all_closing_balances(self.end_date, self.company_id)
        
        def _create_lines(account_types, level):
            account_lines_data = []
            group_total = 0.0
            accounts = Account.search([
                ('account_type', 'in', account_types),
                ('company_id', '=', self.company_id.id)
            ])
            for account in sorted(accounts, key=lambda a: (a.code or '', a.name)):
                balance = closing_balances.get(account.id, 0.0)
                if abs(balance) < 0.01:
                    continue
                account_lines_data.append({
                    'level': level, 'name': f"{'  ' * level}{account.name}",
                    'code': account.code, 'amount': balance,  # Keep raw balance
                    'is_group': False, 'is_total': False,
                })
                group_total += balance
            return account_lines_data, group_total

        # === LIABILITIES SIDE ===
        # Capital Account
        liab_seq += 10
        liab_lines.append({
            'wizard_liab_id': self.id, 'sequence': liab_seq, 'level': 0,
            'name': 'Capital Account', 'is_group': True, 'amount': 0.0,
        })
        capital_idx = len(liab_lines) - 1
        
        account_lines, capital_total = _create_lines(['equity', 'equity_unaffected', 'capital'], 1)
        for line_vals in account_lines:
            liab_seq += 10
            line_vals.update({'wizard_liab_id': self.id, 'sequence': liab_seq})
            liab_lines.append(line_vals)
        
        net_profit_loss = self._get_period_profit_loss(self.start_date, self.end_date)
        liab_seq += 10
        liab_lines.append({
            'wizard_liab_id': self.id, 'sequence': liab_seq, 'level': 1,
            'name': "  Profit & Loss A/c", 'amount': net_profit_loss,
        })
        
        total_capital = capital_total + net_profit_loss
        liab_lines[capital_idx]['amount'] = total_capital

        # Current Liabilities
        liab_seq += 10
        liab_lines.append({
            'wizard_liab_id': self.id, 'sequence': liab_seq, 'level': 0,
            'name': 'Current Liabilities', 'is_group': True, 'amount': 0.0,
        })
        cl_idx = len(liab_lines) - 1
        
        account_lines, cl_total = _create_lines(['liability_payable', 'liability_credit_card', 'liability_current'], 1)
        for line_vals in account_lines:
            liab_seq += 10
            line_vals.update({'wizard_liab_id': self.id, 'sequence': liab_seq})
            liab_lines.append(line_vals)
        liab_lines[cl_idx]['amount'] = cl_total

        # Loans (Liability)
        liab_seq += 10
        liab_lines.append({
            'wizard_liab_id': self.id, 'sequence': liab_seq, 'level': 0,
            'name': 'Loans (Liability)', 'is_group': True, 'amount': 0.0,
        })
        ncl_idx = len(liab_lines) - 1
        
        account_lines, ncl_total = _create_lines(['liability_non_current'], 1)
        for line_vals in account_lines:
            liab_seq += 10
            line_vals.update({'wizard_liab_id': self.id, 'sequence': liab_seq})
            liab_lines.append(line_vals)
        liab_lines[ncl_idx]['amount'] = ncl_total

        total_liabilities = total_capital + cl_total + ncl_total
        
        liab_seq += 10
        liab_lines.append({
            'wizard_liab_id': self.id, 'sequence': liab_seq, 'level': 0,
            'name': 'Total', 'amount': abs(total_liabilities), 'is_total': True
        })

        # === ASSETS SIDE ===
        # Fixed Assets
        asset_seq += 10
        asset_lines.append({
            'wizard_asset_id': self.id, 'sequence': asset_seq, 'level': 0,
            'name': 'Fixed Assets', 'is_group': True, 'amount': 0.0,
        })
        fa_idx = len(asset_lines) - 1
        
        account_lines, fa_total = _create_lines(['asset_fixed', 'asset_non_current'], 1)
        for line_vals in account_lines:
            asset_seq += 10
            line_vals.update({'wizard_asset_id': self.id, 'sequence': asset_seq})
            asset_lines.append(line_vals)
        asset_lines[fa_idx]['amount'] = fa_total

        # Current Assets
        asset_seq += 10
        asset_lines.append({
            'wizard_asset_id': self.id, 'sequence': asset_seq, 'level': 0,
            'name': 'Current Assets', 'is_group': True, 'amount': 0.0,
        })
        ca_idx = len(asset_lines) - 1
        
        account_lines, ca_total = _create_lines(['asset_receivable', 'asset_cash', 'asset_current', 'asset_prepayment'], 1)
        for line_vals in account_lines:
            asset_seq += 10
            line_vals.update({'wizard_asset_id': self.id, 'sequence': asset_seq})
            asset_lines.append(line_vals)
        asset_lines[ca_idx]['amount'] = ca_total

        total_assets = fa_total + ca_total
        
        asset_seq += 10
        asset_lines.append({
            'wizard_asset_id': self.id, 'sequence': asset_seq, 'level': 0,
            'name': 'Total', 'amount': abs(total_assets), 'is_total': True
        })

        self.env['tally.balance.sheet.line'].create(liab_lines)
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
                worksheet.write(row, 1, abs(line.amount), formats['total'])
            elif line.is_group:
                worksheet.write(row, 0, line.name, formats['group'])
                worksheet.write(row, 1, abs(line.amount) if abs(line.amount) > 0.01 else '', formats['group_number'])
            else:
                worksheet.write(row, 0, line.name, formats['account'])
                worksheet.write(row, 1, abs(line.amount) if abs(line.amount) > 0.01 else '', formats['number'])
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
            # Liabilities
            if i < len(self.liability_line_ids):
                line = self.liability_line_ids[i]
                if line.is_total:
                    worksheet.write(row, 0, line.name, formats['total_text'])
                    worksheet.write(row, 1, abs(line.amount), formats['total'])
                elif line.is_group:
                    worksheet.write(row, 0, line.name, formats['group'])
                    worksheet.write(row, 1, abs(line.amount) if abs(line.amount) > 0.01 else '', formats['group_number'])
                else:
                    worksheet.write(row, 0, line.name, formats['account'])
                    worksheet.write(row, 1, abs(line.amount) if abs(line.amount) > 0.01 else '', formats['number'])
            
            # Assets
            if i < len(self.asset_line_ids):
                line = self.asset_line_ids[i]
                if line.is_total:
                    worksheet.write(row, 2, line.name, formats['total_text'])
                    worksheet.write(row, 3, abs(line.amount), formats['total'])
                elif line.is_group:
                    worksheet.write(row, 2, line.name, formats['group'])
                    worksheet.write(row, 3, abs(line.amount) if abs(line.amount) > 0.01 else '', formats['group_number'])
                else:
                    worksheet.write(row, 2, line.name, formats['account'])
                    worksheet.write(row, 3, abs(line.amount) if abs(line.amount) > 0.01 else '', formats['number'])
            
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