{
    'name': 'Accounting Excel Reports',
    'version': '17.0.1.2.0', # Increment version
    'category': 'Accounting',
    'summary': 'Generate Excel and View Tally-style reports for Trial Balance, Balance Sheet, and P&L',
    'description': """
        This module provides Excel export and Odoo view functionality for:
        - Trial Balance (Tally Format)
        - Balance Sheet (Tally Format)
        - Profit & Loss Account (Tally Format)
        Uses hard-coded grouping based on account types for compatibility with Invoicing module.
        
        New in 1.2.0:
        - Added QWeb HTML views for all three reports for a Tally-style on-screen display.
    """,
    'author': 'Concept Solutions ',
    'website': 'https://www.csloman.com',
    'depends': ['account', 'web'], # Added 'web' for QWeb templates
    'data': [
        'security/ir.model.access.csv',
        'wizard/trial_balance_wizard_view.xml',
        'wizard/balance_sheet_wizard_view.xml',
        'wizard/profit_loss_wizard_view.xml',
        'views/accounting_reports_menu.xml',
        'views/tally_report_views.xml',
        'views/tally_report_actions.xml', # --- NEW FILE ---
        'views/tally_report_templates.xml', # --- NEW FILE ---
    ],
    'installable': True,
    'application': False,
    'auto_install': False,
    'license': 'LGPL-3',
}