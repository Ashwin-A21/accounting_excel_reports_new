{
    'name': 'Accounting Excel Reports',
    'version': '17.0.1.3.0', # Incremented version
    'category': 'Accounting',
    'summary': 'Generate Excel and View Tally-style reports for Trial Balance, Balance Sheet, and P&L',
    'description': """
        This module provides Excel export and Odoo view functionality for:
        - Trial Balance (Tally Format)
        - Balance Sheet (Tally Format)
        - Profit & Loss Account (Tally Format)
        
        Uses Odoo's core account.move.line calculations for accuracy
        and Tally-style grouping.
        
        New in 1.3.0:
        - Refactored balance calculation to use account.move.line read_group
          for accuracy and performance, removing manual source document parsing.
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