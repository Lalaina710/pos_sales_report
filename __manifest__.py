{
    'name': 'Rapport Ventes PdV (Analyse)',
    'version': '18.0.1.0.3',
    'category': 'Point of Sale/Reporting',
    'summary': 'Export Excel des ventes PdV : Date, Code, Famille, Designation, PV, Point de vente, Qtes, HT, TTC.',
    'author': 'SOPROMER',
    'license': 'LGPL-3',
    'depends': ['point_of_sale'],
    'data': [
        'security/ir.model.access.csv',
        'wizard/pos_sales_report_wizard_views.xml',
    ],
    'external_dependencies': {
        'python': ['xlsxwriter'],
    },
    'installable': True,
    'application': False,
}
