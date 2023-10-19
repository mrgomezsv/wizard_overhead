# -*- coding: utf-8 -*-

from odoo import api, fields, models
from datetime import datetime
from dateutil import rrule, tz
from odoo.tools import float_compare
import calendar


class TrprovOverheadTr(models.TransientModel):
    _name = 'trprov.overhead.tr'
    _description = 'Reporte Overhead'

    report_from_date = fields.Date(string="Reporte desde", required=True, default=fields.Date.context_today)
    report_to_date = fields.Date(string="Reporte hasta", required=True, default=fields.Date.context_today)
    categ_ids = fields.Many2many('product.category', string="Categoria de los prodcutos")