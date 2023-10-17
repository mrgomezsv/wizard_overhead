# -*- coding: utf-8 -*-

# from odoo import models, fields, api


# class treming_profil_overhead(models.Model):
#     _name = 'treming_profil_overhead.treming_profil_overhead'
#     _description = 'treming_profil_overhead.treming_profil_overhead'

#     name = fields.Char()
#     value = fields.Integer()
#     value2 = fields.Float(compute="_value_pc", store=True)
#     description = fields.Text()
#
#     @api.depends('value')
#     def _value_pc(self):
#         for record in self:
#             record.value2 = float(record.value) / 100
