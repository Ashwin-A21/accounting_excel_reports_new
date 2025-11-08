{
    'name': 'Accounting Excel Reports',
    'version': '17.0.2.2.0', # Incremented version
    'category': 'Accounting',
    'summary': 'Generate Excel and View Tally-style reports for Trial Balance, Balance Sheet, and P&L',
    'description': """
        This module provides Excel export and Odoo view functionality for:
        - Trial Balance (Tally Format)
        - Balance Sheet (Tally Format)
        - Profit & Loss Account (Tally Format)
        
        New in 2.0.0:
        - Implements "Standalone" Tally classification logic.
        - Ignores client's Chart of Account types and classifies based on account names 
          (e.g., 'bank', 'sales', 'sundry creditor') for consistent Tally grouping.
        - Fixes balance calculation to use true net balance (Debit - Credit) 
          for all accounts, ignoring reconciliation status. This fixes bugs
          related to paid/unpaid entries showing incorrect balances.
        - Corrects horizontal Balance Sheet total calculations.
        
        New in 2.2.0:
        - Fixes QWeb report templates (Balance Sheet and P&L) to correctly
          display negative numbers (e.g., for suspense accounts or debit-balance
          liabilities) instead of forcing them to be positive with abs().
    """,
    'author': 'Concept Solutions ',
    'website': 'https://www.csloman.com',
    'depends': ['account', 'web'],
    'data': [
        'security/ir.model.access.csv',
        'wizard/trial_balance_wizard_view.xml',
        'wizard/balance_sheet_wizard_view.xml',
        'wizard/profit_loss_wizard_view.xml',
        'views/accounting_reports_menu.xml',
        'views/tally_report_views.xml',
        'views/tally_report_actions.xml',
        'views/tally_report_templates.xml',
    ],
    'installable': True,
    'application': False,
    'auto_install': False,
    'license': 'LGPL-3',
}