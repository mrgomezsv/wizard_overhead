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
    res_partner_ids = fields.Many2many('res.partner', string="Vendedor")
    month_range = fields.Selection(selection=lambda self: self.trfore_get_month_range_tr(), string="Rango de mes",
                                   default='3', required=True)
    company_id = fields.Many2one(comodel_name="res.company", string="Compañia", required=True,
                                 default=lambda self: self.env.company.id)
    file_content = fields.Binary(string="Archivo Contenido")

    def trfore_get_month_range_tr(self):
        month_range = [
            ('1', '1 mes'),
            ('2', '2 meses'),
            ('3', '3 meses'),
            ('4', '4 meses'),
            ('5', '5 meses'),
            ('6', '6 meses'),
            ('7', '7 meses'),
            ('8', '8 meses'),
            ('10', '10 meses'),
            ('11', '11 meses'),
            ('12', '12 meses'),

        ]
        return month_range
#
#     def calc_new_time(self, date, from_tz=False, to_tz=False):
#         if not from_tz:
#             from_tz = self.env.user.tz
#
#         if not to_tz:
#             to_tz = "UTC"
#
#         from_zone = tz.gettz(from_tz)
#         to_zone = tz.gettz(to_tz)
#         holder = date
#         local_date = holder.replace(tzinfo=from_zone)
#         new_date = local_date.astimezone(to_zone)
#         return new_date
#
#     def action_print_expiry_report_xlsx(self):
#         selected_categories = self.categ_ids.ids
#         domain = [('type', '=', 'product'), ("company_id", "in", [False, self.company_id.id])]
#         if selected_categories:
#             domain.append(('categ_id', 'in', selected_categories))
#
#         products = self.env['product.product'].search(domain, order='default_code desc')
#
#         products_transits = {}
#         purchase_domain = [("product_id", "in", products.ids), ("order_id.state", "=", "purchase"),
#                            ("order_id.company_id", "=", self.company_id.id)]
#         order_lines = self.env["purchase.order.line"].search(purchase_domain)
#         lines_with_diff = order_lines.filtered(lambda x: x.product_qty - x.qty_received)
#
#         if self.res_partner_ids:
#             lines_with_diff = lines_with_diff.filtered(lambda x: x.order_id.partner_id.id in self.res_partner_ids.ids)
#
#         for curr_line in lines_with_diff:
#             target_product = curr_line.product_id
#             products_transits.setdefault(target_product.id, [])
#             products_transits[target_product.id].append(curr_line.order_id.id)
#
#         for curr_item in products_transits:
#             products_transits[curr_item] = list(set(products_transits[curr_item]))
#
#         products = list(products_transits.keys())
#
#         common_format = '%Y-%m-%d'
#         data = {
#             'data': self.read([])[0],
#             'products': products,
#             'start_date': self.report_from_date.strftime(common_format),
#             'end_date': self.report_to_date.strftime(common_format),
#             'month_range': self.month_range,
#             "products_transits": products_transits,
#             "company_id": self.company_id.id
#         }
#         report_template = 'treming_forecast_report.forecast_report_tr'
#         return self.env.ref(report_template).report_action(self, data=data)
#
#
# class TrprexreProductExpiryReportXlsxTr(models.AbstractModel):
#     _name = 'report.treming_forecast_report.forecast_report_tr'
#     _inherit = 'report.report_xlsx.abstract'
#
#     def calc_cell_format_key(self, curr_value):
#         target_color = "red"
#         fc_params = {"precision_digits": 2}
#         if float_compare(curr_value, 1.51, **fc_params) >= 0 and float_compare(curr_value, 2.99, **fc_params) <= 0:
#             target_color = "yellow"
#         elif float_compare(curr_value, 3, **fc_params) >= 0 and float_compare(curr_value, 4.5, **fc_params) <= 0:
#             target_color = "green"
#         elif float_compare(curr_value, 4.51, **fc_params) >= 0:
#             target_color = "light_blue"
#         return target_color
#
#     def calc_cell_format(self, curr_value, format_holder):
#         format_key = self.calc_cell_format_key(curr_value)
#         to_use_format = format_holder[format_key]["format"]
#         return to_use_format
#
#     def generate_xlsx_report(self, workbook, data, move):
#         sheet = workbook.add_worksheet('Reporte Forecast')
#
#         holder_month_range = int(data.get('month_range'))
#         company_id = data["company_id"]
#
#         holder_start_date = data.get('start_date')
#         holder_end_date = data.get('end_date')
#         common_date_format = '%Y-%m-%d'
#
#         start_date = datetime.strptime(holder_start_date, common_date_format)
#         end_date = datetime.strptime(holder_end_date, common_date_format)
#
#         header_format_company = workbook.add_format(
#             {'align': 'center', 'bold': True, 'border': 1, 'bg_color': '#CCE5FF'})
#
#         comm_num_format = "#,##0.00;-#,##0.00"
#         comm_format_items = {'align': 'center', 'num_format': comm_num_format}
#         info_format = workbook.add_format(comm_format_items)
#         simple_format = workbook.add_format({'align': 'center'})
#         date_format = workbook.add_format({'align': 'center', 'num_format': 'dd/mm/yyyy'})
#         format_holder = {"red": {"bg_color": "#FF0000", "format": False},
#                          "yellow": {"bg_color": "#FFFF00", "format": False},
#                          "green": {"bg_color": "#00B050", "format": False},
#                          "light_blue": {"bg_color": "#00B0F0", "format": False}}
#         for curr_key, curr_item in format_holder.items():
#             holder_values = {}
#             holder_values.update(comm_format_items)
#             holder_values["bg_color"] = curr_item["bg_color"]
#             curr_item["format"] = workbook.add_format(holder_values)
#
#         sheet.write('D1', 'Reporte Forecast', header_format_company)
#
#         sheet.freeze_panes(1, 2)
#         sheet.set_column('A1:A1', 15)
#         sheet.set_column('B2:B2', 25)
#         sheet.set_column('C1:AZ1', 15)
#
#         months_es = {
#             1: 'Enero',
#             2: 'Febrero',
#             3: 'Marzo',
#             4: 'Abril',
#             5: 'Mayo',
#             6: 'Junio',
#             7: 'Julio',
#             8: 'Agosto',
#             9: 'Septiembre',
#             10: 'Octubre',
#             11: 'Noviembre',
#             12: 'Diciembre',
#         }
#
#         month_holder = []
#         for current_day in rrule.rrule(rrule.DAILY, dtstart=start_date, until=end_date):
#             holder = current_day.strftime("%Y-%m")
#             month_holder.append(holder)
#
#         unique_month_holder = set(month_holder)
#         months_obj_holder = []
#         for curr_month in unique_month_holder:
#             holder = curr_month + "-01"
#             holder_date = datetime.strptime(holder, "%Y-%m-%d")
#             months_obj_holder.append(holder_date)
#
#         sorted_months = sorted(months_obj_holder)
#
#         holder_format_values = {'bg_color': '#3F5A80', 'align': 'center', 'bold': True, 'color': 'white'}
#         header_format = workbook.add_format(holder_format_values)
#         header_format.set_text_wrap()
#         header_format.set_align('vcenter')
#         header_format_2 = workbook.add_format(
#             {'bg_color': '#2F4B73', 'align': 'center', 'bold': True, 'color': 'white'})
#         header_format_2.set_text_wrap()
#         header_format_2.set_align('vcenter')
#         sheet.write(6, 0, 'Código', header_format_2)
#         sheet.write(6, 1, 'Nombre', header_format_2)
#         header_row = 6  # Fila de encabezado
#         col_index = 2  # Comienza en la tercera columna
#
#         forecast_report_er = self.env["trfore.forecast.report.tr"]
#         month_range_label = forecast_report_er.trfore_get_month_range_tr()
#         dict_month_label = dict(month_range_label)
#
#         # crea los ecabezados los meses
#         for curr_month in sorted_months:
#             sheet.write(header_row, col_index, f'{months_es[curr_month.month]} {curr_month.year}', header_format)
#             col_index += 1
#         sheet.write(6, col_index, 'Promedio Anual', header_format)
#         col_index += 1
#         to_use_ml = dict_month_label[data['month_range']]
#         holder_label = 'Consumo ultimos {}'.format(to_use_ml)
#         sheet.write(6, col_index, holder_label, header_format)
#         col_index += 1
#         holder_label = 'Promedio ultimos {}'.format(to_use_ml)
#         sheet.write(6, col_index, holder_label, header_format)
#         col_index += 1
#         sheet.write(6, col_index, 'Existencia', header_format)
#         col_index += 1
#         sheet.write(6, col_index, 'Existencias en meses', header_format)
#         col_index += 1
#
#         # Transitos
#         prods_by_orders = sorted(data["products_transits"], key=lambda x: len(data["products_transits"][x]),
#                                  reverse=True)
#         most_order_count = []
#         for item in prods_by_orders:
#             most_order_count = data["products_transits"][item]
#             break
#
#         transit_count = 1
#         for item in most_order_count:
#             holder_label = 'Transito {}'.format(transit_count)
#             sheet.write(6, col_index, holder_label, header_format)
#             col_index += 1
#             sheet.write(6, col_index, 'Fecha', header_format)
#             col_index += 1
#             sheet.write(6, col_index, 'Existencia + transito', header_format)
#             col_index += 1
#             transit_count += 1
#
#         sheet.write(6, col_index, 'Cantidad a Pedir', header_format)
#         col_index += 1
#         sheet.write(6, col_index, 'Cantidad a pedir + Existencia + transito', header_format)
#         col_index += 1
#         sheet.write(6, col_index, 'Lead Time', header_format)
#         col_index += 1
#
#         products_records = self.env["product.product"].browse(data["products"])
#         op_domain = [("product_id", "in", data["products"]), ("company_id", "=", data["company_id"])]
#         products_op = self.env["stock.warehouse.orderpoint"].search(op_domain)
#
#         sorted_products = products_records.sorted(key=lambda x: x.name)
#         for index, product_data in enumerate(sorted_products):
#             row = index + 7
#             sheet.write(row, 0, product_data.default_code or "", info_format)
#             sheet.write(row, 1, product_data.name or "", info_format)
#             col_index = 2
#
#             sum_val_months = []
#             for curr_month in sorted_months:
#                 first_day_of_month = curr_month
#                 holder_ld = calendar.monthrange(curr_month.year, curr_month.month)
#                 last_day_of_month = curr_month.replace(day=holder_ld[1], hour=23, minute=59, second=59)
#
#                 utc_fist = forecast_report_er.calc_new_time(first_day_of_month)
#                 utc_last = forecast_report_er.calc_new_time(last_day_of_month)
#
#                 holder_out_move = self.env['stock.move'].search([
#                     ('product_id', '=', product_data['id']),
#                     ('location_id.usage', 'in', ('internal', 'transit')),
#                     ('location_dest_id.usage', 'not in', ('internal', 'transit')),
#                     ('date', '>=', utc_fist),
#                     ('date', '<=', utc_last),
#                     ('state', '=', 'done'),
#                     ('company_id', '=', company_id)
#                 ])
#
#                 # Hasta donde note, los movimientos que son devoluciones
#                 # tienen establecido el campo origin_returned_move_id
#                 holder_in_move = self.env['stock.move'].search([
#                     ('product_id', '=', product_data['id']),
#                     ('location_id.usage', 'not in', ('internal', 'transit')),
#                     ('location_dest_id.usage', 'in', ('internal', 'transit')),
#                     ('date', '>=', utc_fist),
#                     ('date', '<=', utc_last),
#                     ('origin_returned_move_id', '!=', False),
#                     ('state', '=', 'done'),
#                     ('company_id', '=', company_id)
#                 ])
#
#                 out_move = sum(x.product_uom_qty for x in holder_out_move)
#                 in_move = sum(x.product_uom_qty for x in holder_in_move)
#                 holder_total_consumption = out_move - in_move
#                 sum_val_months.append(holder_total_consumption)
#                 total_consumption = round(holder_total_consumption, 2)
#                 sheet.write(row, col_index, total_consumption, info_format)
#                 col_index += 1
#
#             # promedio anual
#             lst_twelve_months = sum_val_months[-12:]
#             res = sum(lst_twelve_months)
#             holder_yearly_average = res / 12
#             yearly_average = round(holder_yearly_average, 2)
#             sheet.write(row, col_index, yearly_average, info_format)
#             col_index += 1
#
#             # consumo de los ultimos meses
#             lst_selected_months = sum_val_months[-holder_month_range:]
#             result_mont = sum(lst_selected_months)
#             round_result_mont = round(result_mont, 2)
#             holder_months_average = result_mont / holder_month_range
#             months_average = round(holder_months_average, 2)
#             sheet.write(row, col_index, round_result_mont, info_format)
#             col_index += 1
#
#             # promedio de los ultimos meses
#             sheet.write(row, col_index, months_average, info_format)
#             col_index += 1
#
#             # existencias
#             holder_existl = product_data.with_company(data["company_id"]).qty_available
#             existl_prod = round(holder_existl, 2)
#             sheet.write(row, col_index, existl_prod, info_format)
#             col_index += 1
#
#             # existencias en meses
#             monthts_existl = 0
#             if result_mont:
#                 holder_monthts_existl = (existl_prod / result_mont) * holder_month_range
#                 monthts_existl = round(holder_monthts_existl, 2)
#             to_use_format = self.calc_cell_format(monthts_existl, format_holder)
#             sheet.write(row, col_index, monthts_existl, to_use_format)
#
#             # transitos
#             col_index += 1
#             transit = 0
#
#             # Al pasarlo desde el wizard hasta aqui las llaves pasan a ser str
#             str_prod_id = str(product_data.id)
#             current_transits = data["products_transits"][str_prod_id]
#             holder_order = self.env["purchase.order"].browse(current_transits)
#             sorted_orders = holder_order.sorted(key=lambda x: x.date_planned, reverse=True)
#             last_stock_transit = 0
#             most_oc_len = len(most_order_count)
#             total_transit = 0
#             for x in range(0, most_oc_len):
#                 target_order = self.env["purchase.order"]
#                 try:
#                     target_order = sorted_orders[x]
#                 except IndexError as err:
#                     pass
#
#                 if target_order:
#                     curr_order = target_order
#                     holder_order_line = curr_order.order_line.filtered(lambda x: x.product_id.id == product_data.id)
#                     order_product_qty = sum(holder_order_line.mapped('product_qty'))
#                     order_qty_received = sum(holder_order_line.mapped('qty_received'))
#
#                     order_line_diff = order_product_qty - order_qty_received
#                     transit += order_product_qty
#                     total_transit += order_line_diff
#                     sheet.write(row, col_index, order_line_diff, info_format)
#                     col_index += 1
#
#                     order_date = curr_order.date_planned
#                     new_order_date = forecast_report_er.calc_new_time(order_date, from_tz="UTC", to_tz=self.env.user.tz)
#                     date_part = new_order_date.date()
#                     sheet.write_datetime(row, col_index, date_part, date_format)
#                     col_index += 1
#
#                     holder_exis_trns = existl_prod + transit
#                     if result_mont:
#                         hol_exis_trns = (holder_exis_trns / result_mont) * holder_month_range
#                     else:
#                         hol_exis_trns = 0
#                     exis_trns = round(hol_exis_trns, 2)
#
#                     to_use_format = self.calc_cell_format(exis_trns, format_holder)
#                     sheet.write(row, col_index, exis_trns, to_use_format)
#                     col_index += 1
#                     last_stock_transit = exis_trns
#                 else:
#                     sheet.write(row, col_index, 0, info_format)
#                     col_index += 1
#
#                     sheet.write(row, col_index, "", simple_format)
#                     col_index += 1
#
#                     to_use_format = self.calc_cell_format(last_stock_transit, format_holder)
#                     sheet.write(row, col_index, last_stock_transit, to_use_format)
#                     col_index += 1
#
#             provider_delay = 0
#             request_qty = 0
#             product_sellers = product_data.variant_seller_ids
#             if product_sellers:
#                 provider_delay = product_sellers[0].delay
#                 provider_delay = provider_delay / 30
#                 provider_delay = round(provider_delay, 2)
#
#             curr_op = products_op.filtered(lambda x: x.product_id.id == product_data["id"])
#             if float_compare(last_stock_transit, provider_delay, precision_digits=2) == -1 and curr_op:
#                 target_op = curr_op[0]
#                 request_qty = target_op.product_max_qty
#                 request_qty = round(request_qty, 2)
#
#             # Cantidad a Pedir
#             sheet.write(row, col_index, request_qty, info_format)
#             col_index += 1
#
#             # transit lleva la acumulacion de todos los transitos
#             # Cantidad a pedir + Existencia + transito
#             holder = 0
#             if result_mont:
#                 holder = existl_prod + total_transit + request_qty
#                 holder = holder / result_mont
#                 holder = holder * holder_month_range
#                 holder = round(holder, 2)
#
#             to_use_format = self.calc_cell_format(holder, format_holder)
#             sheet.write(row, col_index, holder, to_use_format)
#             col_index += 1
#
#             sheet.write(row, col_index, provider_delay, info_format)
#
#         return workbook
