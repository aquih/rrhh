# -*- coding: utf-8 -*-
# Part of Odoo. See LICENSE file for full copyright and licensing details.

from odoo import api, fields, models, _
from odoo.exceptions import UserError
import base64
import xlsxwriter
import logging
import io
import datetime

class rrhh_formato_carga_isr(models.TransientModel):
    _name = 'rrhh.formato_carga_isr'

    formulario = fields.Char('No. Formulario')
    fecha_inicio = fields.Date('Fecha inicio')
    fecha_fin = fields.Date('Fecha fin')
    archivo = fields.Binary('Archivo')
    name =  fields.Char('File Name', size=32)

    def _get_empleados(self, empleados):
        empleado_id = self.env['hr.employee'].search([('id', 'in', empleados)])
        return empleado_id

    def _get_informacion(self, empleados_ids, fecha_inicio, fecha_fin):
        retenciones_sobre_ventas = 0
        impuesto_devuelto_compensado = 0
        impuesto_retenido_pagar = 0
        nomina_id = self.env['hr.payslip'].search([('employee_id','in', empleados_ids),('date_from', '>=', fecha_inicio),('date_to','<=', fecha_fin)])
        if nomina_id:
            logging.warning(nomina_id)
            for nomina in nomina_id:
                if nomina.line_ids:
                    for linea in nomina.line_ids:
                        if linea.salary_rule_id.id in nomina.employee_id.company_id.retenciones_sobre_ventas_ids.ids:
                            retenciones_sobre_ventas += linea.total
                        if linea.salary_rule_id.id in nomina.employee_id.company_id.impuesto_devuelto_compensado_ids.ids:
                            impuesto_devuelto_compensado += linea.total
        impuesto_retenido_pagar = retenciones_sobre_ventas - impuesto_devuelto_compensado
        return {'retenciones_sobre_ventas': retenciones_sobre_ventas,'impuesto_devuelto_compensado': impuesto_devuelto_compensado, 'impuesto_retenido_pagar': impuesto_retenido_pagar}

    def _get_mes_letras(self, fecha):
        en_letras = {
            0: 'Enero',
            1: 'Febrero',
            2: 'Marzo',
            3: 'Abril',
            4: 'Mayo',
            5: 'Junio',
            6: 'Julio',
            7: 'Agosto',
            8: 'Septiembre',
            9: 'Octubre',
            10: 'Noviembre',
            11: 'Diciembre',
        }
        mes = fecha.month
        return en_letras[int(mes) - 1]

    def generar(self):
        for w in self:
            datos = ''
            f = io.BytesIO()
            libro = xlsxwriter.Workbook(f)
            formato_fecha = libro.add_format({'num_format': 'dd/mm/yy'})

            hoja = libro.add_worksheet('SAT-1331 PAGADOS A LA SAT')

            empleados_ids = self._get_empleados(w.env.context.get('active_ids', [])).ids
            informacion = self._get_informacion(empleados_ids, w.fecha_inicio, w.fecha_fin)
            hoja.write(0, 0, 'No. DE FORMULARIO ')
            hoja.write(0, 1, 'PERIODO DE IMPOSICIÓN ')
            hoja.write(0, 2, 'TOTAL DE RETENCIONES SOBRE RENTAS DEL TRABAJO')
            hoja.write(0, 3, 'IMPUESTO DEVUELTO COMPENSADO')
            hoja.write(0, 4, 'IMPUESTO RETENIDO A PAGAR')

            periodo_imposicion = self._get_mes_letras(w.fecha_fin)
            hoja.write(1, 0, w.formulario)
            hoja.write(1, 1, periodo_imposicion)
            hoja.write(1, 2, informacion['retenciones_sobre_ventas'])
            hoja.write(1, 3, informacion['impuesto_devuelto_compensado'])
            hoja.write(1, 4, informacion['impuesto_retenido_pagar'])

            libro.close()
            datos = base64.b64encode(f.getvalue())
            self.write({'archivo': datos, 'name':'ISR Asalariado_SAT 1331.xls'})

        return {
            'context': self.env.context,
            'view_type': 'form',
            'view_mode': 'form',
            'res_model': 'rrhh.formato_carga_isr',
            'res_id': self.id,
            'view_id': False,
            'type': 'ir.actions.act_window',
            'target': 'new',
            }
