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

    def _get_closing_balances(self, date_to, company_id):
        """
        Get closing balances from core journal items - TALLY STANDARD
        Returns balances as NATURAL AMOUNTS (always positive):
        - Assets: Positive (Debit balance)
        - Liabilities: Positive (Credit balance)
        - Capital: Positive (Credit balance)
        """
        balances = defaultdict(float)
        
        # Define Balance Sheet account types
        bs_account_types = [
            'asset_receivable', 'asset_cash', 'asset_current', 'asset_prepayment',
            'asset_fixed', 'asset_non_current',
            'liability_payable', 'liability_current', 'liability_non_current', 
            'liability_credit_card',
            'equity', 'equity_unaffected'
        ]
        
        # Find all relevant accounts
        accounts = self.env['account.account'].search([
            ('account_type', 'in', bs_account_types),
            ('company_id', '=', company_id.id)
        ])
        if not accounts:
            return balances

        # Create a map of account_id to its type
        account_type_map = {acc.id: acc.account_type for acc in accounts}
        
        # Domain for all posted journal items up to the end date
        domain = [
            ('move_id.state', '=', 'posted'),
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

        # Asset types (natural debit balance)
        asset_types = ('asset_receivable', 'asset_cash', 'asset_current', 
                       'asset_fixed', 'asset_non_current', 'asset_prepayment')

        # Calculate the natural balance for each account
        for res in read_group_result:
            if not res['account_id']:
                continue
                
            account_id = res['account_id'][0]
            account_type = account_type_map.get(account_id)
            
            if not account_type:
                continue
                
            debit = res['debit'] or 0.0
            credit = res['credit'] or 0.0
            
            if account_type in asset_types:
                # Assets: Natural balance is Debit - Credit
                balances[account_id] = debit - credit
            else:
                # Liabilities & Equity: Natural balance is Credit - Debit
                balances[account_id] = credit - debit

        return balances

    def _get_period_profit_loss(self, date_from, date_to, company_id):
        """
        Calculate Net Profit/Loss from core journal items - TALLY STANDARD
        Returns: Positive for Profit, Negative for Loss
        """
        
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
            return 0.0

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
        
        income_total = 0.0
        expense_total = 0.0
        
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
                # Income: Credit - Debit
                income_total += (credit - debit)
            else:
                # Expense: Debit - Credit
                expense_total += (debit - credit)

        # Net P/L = Total Income - Total Expense
        # Positive result = Profit, Negative result = Loss
        return income_total - expense_total

    def _prepare_vertical_report_lines(self):
        """Prepare Balance Sheet in vertical format - TALLY STANDARD"""
        self.ensure_one()
        self.line_ids.unlink()
        lines = []
        sequence = 0
        Account = self.env['account.account']
        
        closing_balances = self._get_closing_balances(self.end_date, self.company_id)

        def _create_lines_for_type(account_types, level):
            """Returns lines with POSITIVE balances"""
            account_lines = []
            group_total = 0.0
            accounts = Account.search([
                ('account_type', 'in', account_types),
                ('company_id', '=', self.company_id.id)
            ])
            for account in sorted(accounts, key=lambda a: (a.code or '', a.name)):
                # Use the pre-calculated natural balance
                balance = closing_balances.get(account.id, 0.0)
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

        # === LIABILITIES & CAPITAL ===
        
        # Capital Account
        sequence += 10
        lines.append({
            'wizard_id': self.id, 'sequence': sequence, 'level': 0,
            'name': 'Capital Account', 'is_group': True, 'amount': 0.0,
        })
        capital_group_idx = len(lines) - 1
        
        account_lines, capital_total = _create_lines_for_type(['equity', 'equity_unaffected'], 1)
        for line_vals in account_lines:
            sequence += 10
            line_vals.update({'wizard_id': self.id, 'sequence': sequence})
            lines.append(line_vals)
        
        # Add Current Period P&L to Capital
        net_profit_loss = self._get_period_profit_loss(self.start_date, self.end_date, self.company_id)
        
        if abs(net_profit_loss) > 0.01:
            sequence += 10
            pl_name = "  Profit & Loss A/c (Profit)" if net_profit_loss >= 0 else "  Profit & Loss A/c (Loss)"
            lines.append({
                'wizard_id': self.id, 'sequence': sequence, 'level': 1,
                'name': pl_name, 'amount': abs(net_profit_loss),
            })
        
        # Total Capital (Capital + Profit - Loss)
        # Note: capital_total is already positive (natural balance)
        total_capital = capital_total + net_profit_loss
        lines[capital_group_idx]['amount'] = abs(total_capital)

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

        # Total Liabilities side (Capital + P&L + Liabilities)
        total_liabilities_side = abs(total_capital) + cl_total + ncl_total

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

        # === TOTAL (Should match on both sides) ===
        sequence += 10
        balance_total = max(total_liabilities_side, total_assets)
        lines.append({
            'wizard_id': self.id, 'sequence': sequence, 'level': 0,
            'name': 'Total', 'amount': balance_total,
            'is_total': True,
        })

        self.env['tally.balance.sheet.line'].create(lines)

    def _prepare_horizontal_report_lines(self):
        """Prepare Balance Sheet in horizontal format - TALLY STANDARD"""
        self.ensure_one()
        self.liability_line_ids.unlink()
        self.asset_line_ids.unlink()
        
        liab_lines = []
        asset_lines = []
        liab_seq = 0
        asset_seq = 0
        
        Account = self.env['account.account']
        closing_balances = self._get_closing_balances(self.end_date, self.company_id)
        
        def _create_lines(account_types, level):
            """Returns lines with POSITIVE balances"""
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
                    'level': level, 'name': f"{'  ' * level}{account.name} ({account.code or 'N/A'})",
                    'code': account.code, 'amount': abs(balance),
                    'is_group': False, 'is_total': False,
                })
                group_total += abs(balance)
            return account_lines_data, group_total

        # === LIABILITIES SIDE ===
        
        # Capital Account
        liab_seq += 10
        liab_lines.append({
            'wizard_liab_id': self.id, 'sequence': liab_seq, 'level': 0,
            'name': 'Capital Account', 'is_group': True, 'amount': 0.0,
        })
        capital_idx = len(liab_lines) - 1
        
        account_lines, capital_total = _create_lines(['equity', 'equity_unaffected'], 1)
        for line_vals in account_lines:
            liab_seq += 10
            line_vals.update({'wizard_liab_id': self.id, 'sequence': liab_seq})
            liab_lines.append(line_vals)
        
        net_profit_loss = self._get_period_profit_loss(self.start_date, self.end_date, self.company_id)
        if abs(net_profit_loss) > 0.01:
            liab_seq += 10
            pl_name = "  Profit & Loss A/c (Profit)" if net_profit_loss >= 0 else "  Profit & Loss A/c (Loss)"
            liab_lines.append({
                'wizard_liab_id': self.id, 'sequence': liab_seq, 'level': 1,
                'name': pl_name, 'amount': abs(net_profit_loss),
            })
        
        total_capital = capital_total + net_profit_loss
        liab_lines[capital_idx]['amount'] = abs(total_capital)

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

        total_liabilities = abs(total_capital) + cl_total + ncl_total
        
        liab_seq += 10
        liab_lines.append({
            'wizard_liab_id': self.id, 'sequence': liab_seq, 'level': 0,
            'name': 'Total', 'amount': total_liabilities, 'is_total': True
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
            'name': 'Total', 'amount': total_assets, 'is_total': True
        })
        
        # --- Create Lines ---
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