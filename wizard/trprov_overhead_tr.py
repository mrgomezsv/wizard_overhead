# -*- coding: utf-8 -*-

from odoo import api, fields, models
import openpyxl
from io import BytesIO
import base64

class TrprovOverheadTr(models.TransientModel):
    _name = 'trprov.overhead.tr'
    _description = 'Reporte Overhead'

    report_from_date = fields.Date(string="Reporte desde", required=True, default=fields.Date.context_today)
    report_to_date = fields.Date(string="Reporte hasta", required=True, default=fields.Date.context_today)
    categ_ids = fields.Many2many('product.category', string="Categoria de los productos")
    res_seller_ids = fields.Many2many('account.analytic.account', string="Cuentas Analíticas")
    company_id = fields.Many2one(comodel_name="res.company", string="Compañia", required=True,
                                 default=lambda self: self.env.company.id)
    file_content = fields.Binary(string="Archivo Contenido")

    def action_generate_excel(self):
        wk_book = openpyxl.Workbook()
        wk_sheet = wk_book.active
        wk_sheet.title = "Estado de Resultados por Vendedor"

        # Agrega el título en la celda C2 a S2
        wk_sheet['C2'] = "Estado de Resultado por Vendedor con Overhead"
        wk_sheet.merge_cells('C2:S2')
        title_cell = wk_sheet['C2']
        title_cell.alignment = openpyxl.styles.Alignment(horizontal="center", vertical="center")
        title_cell.font = openpyxl.styles.Font(bold=True, size=14)

        headers = [
            "Nombre Cto. Costo", "Concepto", "Nombre Cta Agrupador", "Nombre Cuenta", "Enero", "Febrero", "Marzo",
            "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
            "Total Resultado"
        ]

        # Comienza a llenar los encabezados a partir de la tercera fila y tercera columna
        row_index = 3
        col_index = 3
        for header in headers:
            cell = wk_sheet.cell(row=row_index, column=col_index)
            cell.value = header
            col_index += 1

        data = self.obtener_datos_estado_resultados()

        for row_index, row_data in enumerate(data, 4):  # Comienza a llenar los datos a partir de la cuarta fila
            vendedor = row_data['vendedor']
            ventas = row_data['ventas']
            costo_ventas = row_data['costo_ventas']
            gastos_operativos = row_data['gastos_operativos']
            utilidad_neta = row_data['utilidad_neta']
            cuenta_analitica = row_data['cuenta_analitica']

            # Llena los valores en las columnas correspondientes
            wk_sheet.cell(row=row_index, column=1, value=vendedor)
            wk_sheet.cell(row=row_index, column=2, value=ventas)
            wk_sheet.cell(row=row_index, column=3, value=costo_ventas)
            wk_sheet.cell(row=row_index, column=4, value=gastos_operativos)
            wk_sheet.cell(row=row_index, column=5, value=utilidad_neta)
            wk_sheet.cell(row=row_index, column=6, value=cuenta_analitica)

        output = BytesIO()
        wk_book.save(output)

        base64_content = base64.b64encode(output.getvalue())

        self.write({'file_content': base64_content})

        return {
            'type': 'ir.actions.act_url',
            'url': "web/content/?model={}&id={}&field=file_content&filename=Reporte_Overhead.xlsx&download=true".format(
                self._name, self.id),
            'target': 'self',
        }

    def action_custom_button(self):
        return self.action_generate_excel()

    def _init_options_buttons(self, options, previous_options=None):
        super(AccountReport, self)._init_options_buttons(options, previous_options=previous_options)
        reporte = self.env.ref(
            'account_reports.profit_and_loss')
        if self.id == reporte.id:
            options['buttons'].append({'name': 'Overhead', 'action': 'action_generate_excel', 'sequence': 100})

    def obtener_datos_estado_resultados(self):
        data = []

        for analytic_account in self.res_seller_ids:
            data.append({
                'vendedor': 'Vendedor 1',  # Agrega lógica para obtener el nombre del vendedor
                'ventas': 10000,  # Agrega lógica para obtener las ventas
                'costo_ventas': 6000,  # Agrega lógica para obtener el costo de ventas
                'gastos_operativos': 3000,  # Agrega lógica para obtener los gastos operativos
                'utilidad_neta': 1000,  # Agrega lógica para obtener la utilidad neta
                'cuenta_analitica': analytic_account.name,  # Agrega la cuenta analítica
            })

        return data
