# -*- coding: utf-8 -*-
{
    'name': "treming_profil_wizard_overhead_2",

    'summary': """
        Estados de resultados por vendedor, con su Overhead incorporado""",

    'description': """
        Estados de resultados por vendedor, con su Overhead incorporado, exportandolo a Excel
    """,

    'author': "Grupo Treming",
    'website': "https://www.treming.com",

    'category': 'Account',
    'version': '0.1',

    'depends': ['base', 'account'],

    'data': [
        'security/ir.model.access.csv',
        'wizard/trprov_overhead_tr.xml'
    ],
}
