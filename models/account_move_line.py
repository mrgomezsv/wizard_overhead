# -*- coding: utf-8 -*-

from odoo import api, fields, models

class AccountMoveLine(models.Model):
    _inherit = 'account.move.line'

    reseller = fields.Char(related="analytic_line_ids.name")