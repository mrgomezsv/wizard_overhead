# -*- coding: utf-8 -*-
{
    'name': "treming_profil_overhead",

    'summary': """
        Estados de resultados por vendedor, con su Overhead incorporado""",

    'description': """
        Estados de resultados por vendedor, con su Overhead incorporado, exportandolo a Excel
    """,

    'author': "Grupo Treming",
    'website': "https://www.treming.com",

    'category': 'Uncategorized',
    'version': '0.1',

    'depends': ['base'],

    'data': [
        # 'security/ir.model.access.csv',
        'views/views.xml',
        'views/templates.xml',
    ],
}
