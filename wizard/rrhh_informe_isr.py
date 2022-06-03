# -*- coding: utf-8 -*-
# Part of Odoo. See LICENSE file for full copyright and licensing details.

from odoo import api, fields, models, _
from odoo.exceptions import UserError
import base64
import xlsxwriter
import logging
import io

class rrhh_informe_isr(models.TransientModel):
    _name = 'rrhh.informe_isr'

    anio = fields.Integer('Año', required=True)
    archivo = fields.Binary('Archivo')
    name =  fields.Char('File Name', size=32)

    def _get_empleados(self, empleados):
        empleado_id = self.env['hr.employee'].search([('id', 'in', empleados)])
        return empleado_id

    def generar(self):
        datos = ''
        f = io.BytesIO()
        libro = xlsxwriter.Workbook(f)
        hoja = libro.add_worksheet('reporte')
        formato_fecha = libro.add_format({'num_format': 'dd/mm/yy'})

        hoja.write(0, 0, 'NIT Empleado')
        hoja.write(0, 1, 'Nombre del Empleado')
        hoja.write(0, 2, 'Fecha de Alta')
        hoja.write(0, 3, 'Renta Patrono Actual')
        hoja.write(0, 4, 'Bono Anual de Trabajadores')
        hoja.write(0, 5, 'Aguinaldo')
        hoja.write(0, 6, 'NIT Otro Patrono 1')
        hoja.write(0, 7, 'Renta Otro Patrono 1')
        hoja.write(0, 8, 'Retencion Otro Patrono 1')
        hoja.write(0, 9, 'NIT Otro Patrono 2')
        hoja.write(0, 10, 'Renta Otro Patrono 2')
        hoja.write(0, 11, 'Retencion Otro Patrono 2')
        hoja.write(0, 12, 'NIT Otro Patrono 3')
        hoja.write(0, 13, 'Renta Otro Patrono 3')
        hoja.write(0, 14, 'Retencion Otro Patrono 3')
        hoja.write(0, 15, 'NIT Otro Patrono 4')
        hoja.write(0, 16, 'Renta Otro Patrono 4')
        hoja.write(0, 17, 'Retencion Otro Patrono 4')
        hoja.write(0, 18, 'NIT Otro Patrono 5')
        hoja.write(0, 19, 'Renta Otro Patrono 5')
        hoja.write(0, 20, 'Retencion Otro Patrono 5')
        hoja.write(0, 21, 'NIT ex patrono 1')
        hoja.write(0, 22, 'Renta Ex Patrono 1')
        hoja.write(0, 23, 'Retencion Ex Patrono 1')
        hoja.write(0, 24, 'NIT ex patrono 2')
        hoja.write(0, 25, 'Renta Ex Patrono 2')
        hoja.write(0, 26, 'Retencion Ex Patrono 2')
        hoja.write(0, 27, 'NIT ex patrono 3')
        hoja.write(0, 28, 'Renta Ex Patrono 3')
        hoja.write(0, 29, 'Retencion Ex Patrono 3')
        hoja.write(0, 30, 'NIT ex patrono 4')
        hoja.write(0, 31, 'Renta Ex Patrono 4')
        hoja.write(0, 32, 'Retencion Ex Patrono 4')
        hoja.write(0, 33, 'NIT ex patrono 5')
        hoja.write(0, 34, 'Renta Ex Patrono 5')
        hoja.write(0, 35, 'Retencion Ex Patrono 5')
        hoja.write(0, 36, 'Otros Ingresos Gravados')
        hoja.write(0, 36, 'Aguinaldo')
        hoja.write(0, 37, 'Bono Anual de Trabajadores')
        hoja.write(0, 39, 'Cuotas IGSS  y Otros Planes de Seguridad Social ')
        hoja.write(0, 40, 'Fecha de actualización de Renta')
        fila = 1

        for empleado in self._get_empleados(self.env.context.get('active_ids', [])):
            hoja.write(fila, 0, empleado.nit if empleado.nit else '')
            hoja.write(fila, 1, empleado.name)
            hoja.write(fila, 2, empleado.contract_id.date_start if empleado.contract_id else '', formato_fecha)
            fila += 1


        libro.close()
        datos = base64.b64encode(f.getvalue())
        self.write({'archivo': datos, 'name':'informe_isr.xls'})
        return {
            'context': self.env.context,
            'view_type': 'form',
            'view_mode': 'form',
            'res_model': 'rrhh.informe_isr',
            'res_id': self.id,
            'view_id': False,
            'type': 'ir.actions.act_window',
            'target': 'new',
            }

    @api.multi
    def print_report(self):
        datas = {'ids': self.env.context.get('active_ids', [])}
        res = self.read(['anio'])
        res = res and res[0] or {}
        datas['form'] = res
        return self.env.ref('rrhh.action_informe_isr').report_action([], data=datas)
