# -*- encoding: utf-8 -*-

from odoo import models, fields, api, _
from odoo.tools.misc import format_date
import time
import base64
import io
import logging
import datetime
from datetime import datetime

class rrhh_igss_wizard(models.TransientModel):
    _name = 'rrhh.igss.wizard'

    def _default_payslip_run(self):
        if len(self.env.context.get('active_ids', [])) > 0:
            nominas = self.env['hr.payslip.run'].search([('id','in',self.env.context.get('active_ids'))])
            return nominas
        else:
            return None

    payslip_run_id = fields.Many2many('hr.payslip.run', string='Payslip run',default=_default_payslip_run)
    archivo = fields.Binary('Archivo')
    name =  fields.Char('File Name', size=32)
    fecha_inicial = fields.Date('Fecha inicial liquidación')
    fecha_final = fields.Date('Fecha final de liquidación')

    def generar(self):
        datos = ''
        for w in self:
            datos += str(w.payslip_run_id[0].slip_ids[0].company_id.version_mensaje) + '|' + str(datetime.today().strftime('%d/%m/%Y')) + '|' + str(w.payslip_run_id[0].slip_ids[0].company_id.numero_patronal) + '|'+ str(datetime.strptime(str(w.payslip_run_id[0].date_start),'%Y-%m-%d').date().strftime('%m')).lstrip('0')+ '|' + str(datetime.strptime(str(w.payslip_run_id[0].date_start),'%Y-%m-%d').date().strftime('%Y')).lstrip('0') + '|' + str(w.payslip_run_id[0].slip_ids[0].company_id.name) + '|' +str(w.payslip_run_id[0].slip_ids[0].company_id.vat) + '|'+ str(w.payslip_run_id[0].slip_ids[0].company_id.email) + '|' + str(w.payslip_run_id[0].slip_ids[0].company_id.tipo_planilla) + '\r\n'
            datos += '[centros]' + '\r\n'
            for centro in w.payslip_run_id[0].slip_ids[0].company_id.centro_trabajo_ids:
                datos += str(centro.codigo) + '|' + str(centro.nombre) + '|' + str(centro.direccion) + '|' + str(centro.zona) + '|' + str(centro.telefono) + '|' + str(centro.fax) + '|' + str(centro.nombre_contacto) + '|' + str(centro.correo_electronico) + '|' + str(centro.codigo_departamento) + '|' + str(centro.codigo_municipio) + '|' + str(centro.codigo_actividad_economica) + '\r\n'
            datos += '[tiposplanilla]' + '\r\n'
            for tipo_planilla in w.payslip_run_id[0].slip_ids[0].company_id.tipo_planilla_ids:
                datos += str(tipo_planilla.codigo) + '|' + str(tipo_planilla.name) + '|' + str(tipo_planilla.tipo_afiliado) + '|' + str(tipo_planilla.periodo_planilla) + '|' + str(tipo_planilla.codigo_departamento) + '|' + str(tipo_planilla.codigo_actividad_economica) + '|' + str(tipo_planilla.clase_planilla) + '|' + str(tipo_planilla.tiempo_contrato) + '|' + '\r\n'
            datos += '[liquidaciones]' + '\r\n'
            for liquidacion in w.payslip_run_id[0].slip_ids[0].company_id.tipo_planilla_ids.liquidaciones_ids:
                fecha_inicial = format_date(self.env, liquidacion.fecha_inicial, date_format="d/M/y")
                fecha_final = format_date(self.env, liquidacion.fecha_final, date_format="d/M/y")
                datos += str(liquidacion.numero) + '|' + str(liquidacion.tipo_planilla_id.codigo) + '|' + str(fecha_inicial) + '|' + str(fecha_final) + '|' + str(liquidacion.complementaria_original) + '|' + str(liquidacion.numero_nota_cargo) + '\r\n'
            datos += '[empleados]' + '\r\n'
            empleados = {}
            suspensiones = []
            for payslip_run in w.payslip_run_id:
                for slip in payslip_run.slip_ids:
                    if slip.contract_id:
                        if slip.employee_id.id not in empleados:
                            empleados[slip.employee_id.id] = {'empleado_id': slip.employee_id.id,'informacion': [0] * 19,'suspension': ''}

                        contrato_ids = self.env['hr.contract'].search( [['employee_id', '=', slip.employee_id.id]],offset=0,limit=1,order='date_start desc')
                        numero_liquidacion = str(slip.employee_id.numero_liquidacion) if slip.employee_id.numero_liquidacion else ''
                        numero_afiliado = str(slip.employee_id.igss) if slip.employee_id.igss else ''
                        primer_nombre = str(slip.employee_id.primer_nombre) if slip.employee_id.primer_nombre else ''
                        segundo_nombre = str(slip.employee_id.segundo_nombre) if slip.employee_id.segundo_nombre else ''
                        primer_apellido = str(slip.employee_id.primer_apellido) if slip.employee_id.primer_apellido else ''
                        segundo_apellido = str(slip.employee_id.segundo_apellido) if slip.employee_id.segundo_apellido else ''
                        apellido_casada = str(slip.employee_id.apellido_casada) if slip.employee_id.apellido_casada else ''
                        tipo_salario = slip.employee_id.tipo_salario if slip.employee_id.tipo_salario else ''
                        horas_laboradas = ''
                        tiempo_contrato = slip.employee_id.tiempo_contrato if slip.employee_id.tiempo_contrato else ''
                        dias_laborados = 0
                        for linea in slip.worked_days_line_ids:
                            if linea.work_entry_type_id.code == slip.employee_id.company_id.igss_dias_trabajo:
                                dias_laborados = linea.number_of_days
                        sueldo = 0
                        for linea in slip.line_ids:
                            if linea.salary_rule_id.id in slip.employee_id.company_id.sueldo_igss_ids.ids:
                                sueldo += linea.total

                        mes_inicio_contrato = datetime.strptime(str(slip.contract_id.date_start), '%Y-%m-%d').month
                        anio_inicio_contrato = datetime.strptime(str(slip.contract_id.date_start), '%Y-%m-%d').year
                        mes_final_contrato = datetime.strptime(str(slip.contract_id.date_end), '%Y-%m-%d').month if slip.contract_id.date_end else ''
                        anio_final_contrato = datetime.strptime(str(slip.contract_id.date_end), '%Y-%m-%d').year if slip.contract_id.date_end else ''
                        mes_planilla = datetime.strptime(str(payslip_run.date_start), '%Y-%m-%d').month
                        anio_planilla = datetime.strptime(str(payslip_run.date_start), '%Y-%m-%d').year
                        fecha_alta = str(datetime.strptime(str(slip.contract_id.date_start),'%Y-%m-%d').date().strftime('%d/%m/%Y')) if (mes_inicio_contrato == mes_planilla and anio_inicio_contrato == anio_planilla) else ''
                        fecha_baja = str(datetime.strptime(str(slip.contract_id.date_end),'%Y-%m-%d').date().strftime('%d/%m/%Y')) if (mes_final_contrato == mes_planilla and anio_final_contrato == anio_planilla) else ''

                        centro_trabajo = str(slip.employee_id.codigo_centro_trabajo) if slip.employee_id.codigo_centro_trabajo else ''
                        nit = str(slip.employee_id.nit) if slip.employee_id.nit else ''
                        codigo_ocupacion = str(slip.employee_id.codigo_ocupacion) if slip.employee_id.codigo_ocupacion else ''
                        condicion_laboral = str(slip.employee_id.condicion_laboral) if slip.employee_id.condicion_laboral else ''
                        deducciones = ''

                        empleados[slip.employee_id.id]['informacion'][0] = (numero_liquidacion)
                        empleados[slip.employee_id.id]['informacion'][1] = (numero_afiliado)
                        empleados[slip.employee_id.id]['informacion'][2] = (primer_nombre)
                        empleados[slip.employee_id.id]['informacion'][3] = (segundo_nombre)
                        empleados[slip.employee_id.id]['informacion'][4] = (primer_apellido)
                        empleados[slip.employee_id.id]['informacion'][5] = (segundo_apellido)
                        empleados[slip.employee_id.id]['informacion'][6] = (apellido_casada)
                        empleados[slip.employee_id.id]['informacion'][7] += round(sueldo,2)
                        empleados[slip.employee_id.id]['informacion'][8] = (fecha_alta)
                        empleados[slip.employee_id.id]['informacion'][9] = (fecha_baja)
                        empleados[slip.employee_id.id]['informacion'][10] = (centro_trabajo)
                        empleados[slip.employee_id.id]['informacion'][11] = (nit)
                        empleados[slip.employee_id.id]['informacion'][12] = (codigo_ocupacion)
                        empleados[slip.employee_id.id]['informacion'][13] = (condicion_laboral)
                        empleados[slip.employee_id.id]['informacion'][14] = (deducciones)
                        empleados[slip.employee_id.id]['informacion'][15] = (tipo_salario)
                        empleados[slip.employee_id.id]['informacion'][16] = (horas_laboradas)
                        empleados[slip.employee_id.id]['informacion'][17] = (tiempo_contrato)
                        empleados[slip.employee_id.id]['informacion'][18] = int(dias_laborados)

            if empleados:
                for empleado in empleados.values():
                    for dato in empleado['informacion']:
                        index = empleado['informacion'].index(dato)
                        if index != 18:
                            datos += str(dato) + '|'
                        else:
                            datos += str(dato)
                    datos += '\r\n'

                    ausencias = self.env['hr.leave'].search([('employee_id','=', empleado['empleado_id']),('request_date_from','>=',self.fecha_inicial),('request_date_to','<=',self.fecha_final),('state','=','validate')])
                    if ausencias:
                        for ausencia in ausencias:
                            if ausencia.holiday_status_id.suspension_igss:
                                fecha_inicio = str(datetime.strptime(str(ausencia.date_from),'%Y-%m-%d %H:%M:%S').date().strftime('%d/%m/%Y'))
                                fecha_fin = str(datetime.strptime(str(ausencia.date_to),'%Y-%m-%d %H:%M:%S').date().strftime('%d/%m/%Y'))
                                igss = ausencia.employee_id.igss if ausencia.employee_id.igss else ""
                                primer_nombre = ausencia.employee_id.primer_nombre if ausencia.employee_id.primer_nombre else ""
                                segundo_nombre = ausencia.employee_id.segundo_nombre if ausencia.employee_id.segundo_nombre else ""
                                primer_apellido = ausencia.employee_id.primer_apellido if ausencia.employee_id.primer_apellido else ""
                                segundo_apellido = ausencia.employee_id.segundo_apellido if ausencia.employee_id.segundo_apellido else ""
                                apellido_casada = ausencia.employee_id.apellido_casada if ausencia.employee_id.apellido_casada else ""
                                suspensiones.append(numero_liquidacion + '|' + igss + '|' + primer_nombre + '|' + segundo_nombre + '|' + primer_apellido + '|' + segundo_apellido + '|' + apellido_casada  + '|' + str(ausencia.request_date_from.strftime('%d/%m/%Y')) + '|' + str(ausencia.request_date_to.strftime('%d/%m/%Y')) + '|' + '\r\n')

            datos += '[suspendidos]' + '\r\n'
            if suspensiones:
                for suspension in suspensiones:
                    datos += suspension
            datos += '[licencias]' + '\r\n'
            datos += '[juramento]' + '\r\n'
            datos += 'BAJO MI EXCLUSIVA Y ABSOLUTA RESPONSABILIDAD, DECLARO QUE LA INFORMACION QUE AQUI CONSIGNO ES FIEL Y EXACTA, QUE ESTA PLANILLA INCLUYE A TODOS LOS TRABAJADORES QUE ESTUVIERON A MI SERVICIO Y QUE SUS SALARIOS SON LOS EFECTIVAMENTE DEVENGADOS, DURANTE EL MES ARRIBA INDICADO' + '\r\n'
            datos += '[finplanilla]' + '\r\n'
            datos = datos.replace('False', '')
        datos = base64.b64encode(datos.encode("utf-8"))
        self.write({'archivo': datos, 'name':'planilla.txt'})

        return {
            'view_type': 'form',
            'view_mode': 'form',
            'res_model': 'rrhh.igss.wizard',
            'res_id': self.id,
            'view_id': False,
            'type': 'ir.actions.act_window',
            'target': 'new',
        }
