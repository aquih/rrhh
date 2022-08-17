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

    fecha_inicio = fields.Date('Fecha inicio')
    fecha_fin = fields.Date('Fecha fin')
    archivo = fields.Binary('Archivo')
    name =  fields.Char('File Name', size=32)

    def _get_empleados(self, empleados):
        empleado_id = self.env['hr.employee'].search([('id', 'in', empleados)])
        return empleado_id

    def _get_informacion(self, empleados_ids, fecha_inicio, fecha_fin):
        retenciones_sobre_rentas = 0
        impuesto_devuelto_compensado = 0
        impuesto_retenido_pagar = 0
        informacion_mes = {}
        nomina_id = self.env['hr.payslip'].search([('employee_id','in', empleados_ids),('date_from', '>=', fecha_inicio),('date_to','<=', fecha_fin)] , order='date_to asc')
        if nomina_id:
            for nomina in nomina_id:
                mes = (nomina.date_to.month - 1)
                if mes not in informacion_mes:
                    informacion_mes[mes] = {'periodo': self._get_mes_letras(mes),'numero_formulario': nomina.payslip_run_id.formulario,'retenciones_sobre_rentas': 0 ,'impuesto_devuelto_compensado': nomina.payslip_run_id.total_devolucion_isr}

                if nomina.line_ids:
                    for linea in nomina.line_ids:
                        if linea.salary_rule_id.id in nomina.employee_id.company_id.retenciones_sobre_rentas_ids.ids:
                            informacion_mes[mes]['retenciones_sobre_rentas'] += abs(linea.total)
        return informacion_mes

    def _get_mes_letras(self, mes):
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
        return en_letras[int(mes)]

    def generar(self):
        for w in self:
            datos = ''
            f = io.BytesIO()
            libro = xlsxwriter.Workbook(f)
            formato_fecha = libro.add_format({'num_format': 'dd/mm/yy'})

            hoja = libro.add_worksheet('SAT-1331 PAGADOS A LA SAT')

            empleados_ids = self._get_empleados(w.env.context.get('active_ids', [])).ids
            informacion_isr = self._get_informacion(empleados_ids, w.fecha_inicio, w.fecha_fin)
            hoja.write(0, 0, 'No. DE FORMULARIO ')
            hoja.write(0, 1, 'PERIODO DE IMPOSICIÓN ')
            hoja.write(0, 2, 'TOTAL DE RETENCIONES SOBRE RENTAS DEL TRABAJO')
            hoja.write(0, 3, 'IMPUESTO DEVUELTO COMPENSADO')
            hoja.write(0, 4, 'IMPUESTO RETENIDO A PAGAR')

            # periodo_imposicion = self._get_mes_letras(w.fecha_fin)

            if len(informacion_isr) > 0:
                fila = 1
                for mes in informacion_isr:
                    hoja.write(fila, 0, informacion_isr[mes]['numero_formulario'])
                    hoja.write(fila, 1, informacion_isr[mes]['periodo'])
                    hoja.write(fila, 2, informacion_isr[mes]['retenciones_sobre_rentas'])
                    hoja.write(fila, 3, informacion_isr[mes]['impuesto_devuelto_compensado'])
                    impuesto_retenido_pagar = informacion_isr[mes]['retenciones_sobre_rentas'] - informacion_isr[mes]['impuesto_devuelto_compensado']
                    hoja.write(fila, 4, impuesto_retenido_pagar)
                    fila += 1

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
