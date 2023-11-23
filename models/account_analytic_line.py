# -*- coding: utf-8 -*-

from odoo import api, fields, models

class AccountAnalyticLine(models.Model):
    _inherit = 'account.analytic.line'

    trprovwi_general_account_type_tr = fields.Selection(related="general_account_id.account_type", string="Tipo de cuenta financiera", store=True)