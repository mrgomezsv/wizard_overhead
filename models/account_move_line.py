# -*- coding: utf-8 -*-

from odoo import api, fields, models

class AccountMoveLine(models.Model):
    _inherit = 'account.move.line'

    reseller = fields.Char(related="analytic_line_ids.name")
    analytic_boo = fields.Boolean(compute="analytic_bool", store=True)
    #analytic_resseller = fields.Many2many('')

    def analytic_bool(self):
        for rec in self:
            rec.analytic_boo = bool(rec.analytic_distribution)