# -*- coding: utf-8 -*-

from odoo import api, fields, models

class AccountAnalyticAccount(models.Model):
    _inherit = 'account.analytic.account'

    check = fields.Boolean(string="Jorge", help="Campo booleano para la funcionalidad de Jorge")