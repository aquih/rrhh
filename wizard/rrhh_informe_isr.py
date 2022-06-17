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

    def _get_informacion(self, empleado_id, fecha_inicio, fecha_fin):
        otros_ingresos = 0
        viaticos = 0
        igss_total = 0
        aguinaldo_anual = 0
        bono_anual = 0
        renta_patrono_actual = 0
        nomina_id = self.env['hr.payslip'].search([('employee_id','=', empleado_id),('date_from', '>=', fecha_inicio),('date_to','<=',fecha_fin )])
        if nomina_id:
            for nomina in nomina_id:
                if nomina.line_ids:
                    for linea in nomina.line_ids:
                        if linea.salary_rule_id.id in nomina.employee_id.company_id.otros_ingresos_gravados_ids.ids:
                            otros_ingresos += linea.total
                        if linea.salary_rule_id.id in nomina.employee_id.company_id.viaticos_ids.ids:
                            viaticos += linea.total
                        if linea.salary_rule_id.id in nomina.employee_id.company_id.igss_ids.ids:
                            igss_total += linea.total
                        if linea.salary_rule_id.id in nomina.employee_id.company_id.aguinaldo_ids.ids:
                            aguinaldo_anual += linea.total
                        if linea.salary_rule_id.id in nomina.employee_id.company_id.bono_ids.ids:
                            bono_anual += linea.total
                        if linea.salary_rule_id.id in nomina.employee_id.company_id.renta_patrono_actual_ids.ids:
                            renta_patrono_actual += linea.total

        return {'renta_patrono_actual': renta_patrono_actual,'otro_ingreso': otros_ingresos, 'viaticos': viaticos, 'igss_total': igss_total, 'bono_anual': bono_anual ,'aguinaldo_anual': aguinaldo_anual}

    def generar(self):
        datos = ''
        f = io.BytesIO()
        libro = xlsxwriter.Workbook(f)
        hoja = libro.add_worksheet('Cargaproyeccionesyactualización')
        hoja_carga_ajuste = libro.add_worksheet('CargasAjustesysuspensiones')
        hoja_fin_labores = libro.add_worksheet('CargaLiquidación Fin de labores')
        hoja_fin_periodo = libro.add_worksheet('CargaLiquidación Fin Período')
        hoja_retencion = libro.add_worksheet('Retención por pago')


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
        hoja.write(0, 37, 'Aguinaldo')
        hoja.write(0, 38, 'Bono Anual de Trabajadores')
        hoja.write(0, 39, 'Cuotas IGSS  y Otros Planes de Seguridad Social ')
        hoja.write(0, 40, 'Fecha de actualización de Renta')
        fila = 1

        for empleado in self._get_empleados(self.env.context.get('active_ids', [])):
            hoja.write(fila, 0, empleado.nit if empleado.nit else '')
            hoja.write(fila, 1, empleado.name)
            hoja.write(fila, 2, empleado.contract_id.date_start if empleado.contract_id else '', formato_fecha)
            hoja.write(fila, 3, (empleado.contract_id.wage*250) * 12)
            hoja.write(fila, 4, empleado.contract_id.wage)
            hoja.write(fila, 5, empleado.contract_id.wage)

            otra_info = self._get_informacion(empleado.id, '01-07-'+str(self.anio-1), '01-06-'+str(self.anio))
            hoja.write(fila, 36, empleado.contract_id.base_extra * 12)
            hoja.write(fila, 37, empleado.contract_id.wage)
            hoja.write(fila, 38, empleado.contract_id.wage)
            cuota_igss = (empleado.contract_id.wage * 12)*0.0483
            hoja.write(fila, 39, cuota_igss)
            fila += 1

        hoja_carga_ajuste.write(0, 0, 'NIT Empleado')
        hoja_carga_ajuste.write(0, 1, 'AJUSTE/SUSPENSION')
        # hoja_carga_ajuste.write(1, 0, 'GENERAL')
        # hoja_carga_ajuste.write(1, 1, 'GENERAL')
        # hoja_carga_ajuste.write(2, 0, 'sin decimales')
        # hoja_carga_ajuste.write(2, 1, '2 decimales')
        # hoja_carga_ajuste.write(3, 0, 'NIT del empleado sin guión')
        # hoja_carga_ajuste.write(3, 1, 'Retención a efectuar al empleado en el mes seleccionado')
        # hoja_carga_ajuste.write(4, 0, '29532760')
        # hoja_carga_ajuste.write(4, 1, '25.5')

        hoja_fin_labores.write(0, 0, 'NIT empleado')
        hoja_fin_labores.write(0, 1, 'Renta Patrono Actual')
        hoja_fin_labores.write(0, 2, 'Bono Anual de trabajadores (14)')
        hoja_fin_labores.write(0, 3, 'Aguinaldo')
        hoja_fin_labores.write(0, 4, 'NIT Otro Patrono 1')
        hoja_fin_labores.write(0, 5, 'Renta Otro Patrono 1')
        hoja_fin_labores.write(0, 6, 'Retencion Otro Patrono 1')
        hoja_fin_labores.write(0, 7, 'NIT Otro Patrono 2')
        hoja_fin_labores.write(0, 8, 'Renta Otro Patrono 2')
        hoja_fin_labores.write(0, 9, 'Retencion Otro Patrono 2')
        hoja_fin_labores.write(0, 10, 'NIT Otro Patrono 3')
        hoja_fin_labores.write(0, 11, 'Renta Otro Patrono 3')
        hoja_fin_labores.write(0, 12, 'Retencion Otro Patrono 3')
        hoja_fin_labores.write(0, 13, 'NIT Otro Patrono 4')
        hoja_fin_labores.write(0, 14, 'Renta Otro Patrono 4')
        hoja_fin_labores.write(0, 15, 'Retencion Otro Patrono 4')
        hoja_fin_labores.write(0, 16, 'NIT Otro Patrono 5')
        hoja_fin_labores.write(0, 17, 'Renta Otro Patrono 5')
        hoja_fin_labores.write(0, 18, 'Retencion Otro Patrono 5')
        hoja_fin_labores.write(0, 19, 'NIT ex patrono 1')
        hoja_fin_labores.write(0, 20, 'Renta Ex Patrono 1')
        hoja_fin_labores.write(0, 21, 'Retencion Ex Patrono 1')
        hoja_fin_labores.write(0, 22, 'NIT ex patrono 2')
        hoja_fin_labores.write(0, 23, 'Renta Ex Patrono 2')
        hoja_fin_labores.write(0, 24, 'Retencion Ex Patrono 2')
        hoja_fin_labores.write(0, 25, 'NIT ex patrono 3')
        hoja_fin_labores.write(0, 26, 'Renta Ex Patrono 3')
        hoja_fin_labores.write(0, 27, 'Retencion Ex Patrono 3')
        hoja_fin_labores.write(0, 28, 'NIT ex patrono 4')
        hoja_fin_labores.write(0, 29, 'Renta Ex Patrono 4')
        hoja_fin_labores.write(0, 30, 'Retencion Ex Patrono 4')
        hoja_fin_labores.write(0, 31, 'NIT ex patrono 5')
        hoja_fin_labores.write(0, 32, 'Renta Ex Patrono 5')
        hoja_fin_labores.write(0, 33, 'Retencion Ex Patrono 5')
        hoja_fin_labores.write(0, 34, 'Otros ingresos Gravados y Exentos obtenidos en el período')
        hoja_fin_labores.write(0, 35, 'Indemnizaciones o pensiones por causa de muerte')
        hoja_fin_labores.write(0, 36, 'Indemnizaciones por tiempo servido')
        hoja_fin_labores.write(0, 37, 'Remuneraciones de los diplomáticos')
        hoja_fin_labores.write(0, 38, 'Gastos de representación y viáticos comprobables')
        hoja_fin_labores.write(0, 39, 'Aguinaldo')
        hoja_fin_labores.write(0, 40, 'Bono Anual de trabajadores (14)')
        hoja_fin_labores.write(0, 41, 'Cuotas IGSS  y Otros planes de seguridad social')
        hoja_fin_labores.write(0, 42, 'Fecha de Fin de Labores')
        hoja_fin_labores.write(0, 43, 'Ultima Retención')

        fila = 1

        for empleado in self._get_empleados(self.env.context.get('active_ids', [])):
            if empleado.contract_id.date_end:
                hoja_fin_labores.write(fila, 0, empleado.nit if empleado.nit else '')
                hoja_fin_labores.write(fila, 3, (empleado.contract_id.wage))
                hoja_fin_labores.write(fila, 4, empleado.contract_id.wage)
                hoja_fin_labores.write(fila, 5, empleado.contract_id.wage)

                otra_info = self._get_informacion(empleado.id, '01-01-'+str(self.anio), empleado.contract_id.date_end)
                hoja_fin_labores.write(fila, 1, otra_info['renta_patrono_actual'])
                hoja_fin_labores.write(fila, 34, otra_info['otro_ingreso'])


                hoja_fin_labores.write(fila, 38, otra_info['viaticos'])
                hoja_fin_labores.write(fila, 39, empleado.contract_id.wage)
                hoja_fin_labores.write(fila, 40, empleado.contract_id.wage)
                hoja_fin_labores.write(fila, 41, otra_info['igss_total'])
                hoja_fin_labores.write(fila, 42,  str(empleado.contract_id.date_end))
                fila += 1


        hoja_fin_periodo.write(0, 0, 'NIT empleado')
        hoja_fin_periodo.write(0, 1, 'Renta Patrono Actual')
        hoja_fin_periodo.write(0, 2, 'Bono Anual de trabajadores (14)')
        hoja_fin_periodo.write(0, 3, 'Aguinaldo')
        hoja_fin_periodo.write(0, 4, 'NIT Otro Patrono 1')
        hoja_fin_periodo.write(0, 5, 'Renta Otro Patrono 1')
        hoja_fin_periodo.write(0, 6, 'Retencion Otro Patrono 1')
        hoja_fin_periodo.write(0, 7, 'NIT Otro Patrono 2')
        hoja_fin_periodo.write(0, 8, 'Renta Otro Patrono 2')
        hoja_fin_periodo.write(0, 9, 'Retencion Otro Patrono 2')
        hoja_fin_periodo.write(0, 10, 'NIT Otro Patrono 3')
        hoja_fin_periodo.write(0, 11, 'Renta Otro Patrono 3')
        hoja_fin_periodo.write(0, 12, 'Retencion Otro Patrono 3')
        hoja_fin_periodo.write(0, 13, 'NIT Otro Patrono 4')
        hoja_fin_periodo.write(0, 14, 'Renta Otro Patrono 4')
        hoja_fin_periodo.write(0, 15, 'Retencion Otro Patrono 4')
        hoja_fin_periodo.write(0, 16, 'NIT Otro Patrono 5')
        hoja_fin_periodo.write(0, 17, 'Renta Otro Patrono 5')
        hoja_fin_periodo.write(0, 18, 'Retencion Otro Patrono 5')
        hoja_fin_periodo.write(0, 19, 'NIT ex patrono 1')
        hoja_fin_periodo.write(0, 20, 'Renta Ex Patrono 1')
        hoja_fin_periodo.write(0, 21, 'Retencion Ex Patrono 1')
        hoja_fin_periodo.write(0, 22, 'NIT ex patrono 2')
        hoja_fin_periodo.write(0, 23, 'Renta Ex Patrono 2')
        hoja_fin_periodo.write(0, 24, 'Retencion Ex Patrono 2')
        hoja_fin_periodo.write(0, 25, 'NIT ex patrono 3')
        hoja_fin_periodo.write(0, 26, 'Renta Ex Patrono 3')
        hoja_fin_periodo.write(0, 27, 'Retencion Ex Patrono 3')
        hoja_fin_periodo.write(0, 28, 'NIT ex patrono 4')
        hoja_fin_periodo.write(0, 29, 'Renta Ex Patrono 4')
        hoja_fin_periodo.write(0, 30, 'Retencion Ex Patrono 4')
        hoja_fin_periodo.write(0, 31, 'NIT ex patrono 5')
        hoja_fin_periodo.write(0, 32, 'Renta Ex Patrono 5')
        hoja_fin_periodo.write(0, 33, 'Retencion Ex Patrono 5')
        hoja_fin_periodo.write(0, 34, 'Otros ingresos Gravados y Exentos obtenidos en el período')
        hoja_fin_periodo.write(0, 35, 'Indemnizaciones o pensiones por causa de muerte')
        hoja_fin_periodo.write(0, 36, 'Indemnizaciones por tiempo servido')
        hoja_fin_periodo.write(0, 37, 'Remuneraciones de los diplomáticos')
        hoja_fin_periodo.write(0, 38, 'Gastos de representación y viáticos comprobables')
        hoja_fin_periodo.write(0, 39, 'Aguinaldo')
        hoja_fin_periodo.write(0, 40, 'Bono Anual de trabajadores (14)')
        hoja_fin_periodo.write(0, 41, 'Cuotas IGSS  y Otros planes de seguridad social')
        hoja_fin_periodo.write(0, 42, 'Seguros')
        hoja_fin_periodo.write(0, 43, 'Planilla')
        hoja_fin_periodo.write(0, 43, 'Otras Donaciones')
        fila = 1

        for empleado in self._get_empleados(self.env.context.get('active_ids', [])):
            otra_info = self._get_informacion(empleado.id, '01-07-'+str(self.anio-1), '31-12-'+str(self.anio-1))
            hoja_fin_labores.write(fila, 0, empleado.nit if empleado.nit else '')
            hoja_fin_labores.write(fila, 1, (empleado.contract_id.wage*2))
            hoja_fin_labores.write(fila, 2, otra_info['bono_anual'])
            hoja_fin_labores.write(fila, 3, otra_info['aguinaldo_anual'])


            otra_info = self._get_informacion(empleado.id, '01-07-'+str(self.anio), empleado.contract_id.date_end)

            hoja_fin_labores.write(fila, 34, otra_info['otro_ingreso'])


            hoja_fin_labores.write(fila, 38, otra_info['viaticos'])
            hoja_fin_labores.write(fila, 39, empleado.contract_id.wage)
            hoja_fin_labores.write(fila, 40, empleado.contract_id.wage)
            hoja_fin_labores.write(fila, 41, otra_info['igss_total'])
            fila += 1


        hoja_retencion.write(0, 0, 'NIT empleado')
        hoja_retencion.write(0, 1, 'Base Gravada o Pagada')
        hoja_retencion.write(0, 2, 'Retencion por pago')
        hoja_retencion.write(0, 3, 'Fecha de retencion')


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
