from odoo import models, fields, api, _
from odoo.tools.misc import format_date
import base64
import logging
from datetime import datetime

class rrhh_igss_wizard(models.TransientModel):
    _name = 'rrhh.archivo_igss.wizard'
    _description = 'Wizard para generar archivo de IGSS'

    def _default_payslip_run(self):
        if len(self.env.context.get('active_ids', [])) > 0:
            nominas = self.env['hr.payslip.run'].search([('id','in',self.env.context.get('active_ids'))])
            return nominas
        else:
            return None

    payslip_run_id = fields.Many2many('hr.payslip.run', string='Payslip run', default=_default_payslip_run)
    archivo = fields.Binary('Archivo')
    name =  fields.Char('File Name', size=32)
    fecha_inicial = fields.Date('Fecha inicial liquidación')
    fecha_final = fields.Date('Fecha final de liquidación')

    def print_report_excel(self):
        datos = ''
        for w in self:
            lote_id = w.payslip_run_id[0]
            compania_id = lote_id.company_id

            datos += str(compania_id.version_mensaje) + '|' + str(format_date(self.env, datetime.today(), date_format="dd/MM/yyyy")) + '|' + str(compania_id.numero_patronal) + '|'+ str(format_date(self.env, lote_id.date_start, date_format="MM")) + '|' + str(format_date(self.env, lote_id.date_start, date_format="YYYY")) + '|' + str(compania_id.name) + '|' +str(compania_id.vat) + '|'+ str(compania_id.email) + '|' + str(compania_id.tipo_planilla) + '\r\n'
            datos += '[centros]' + '\r\n'
            for centro in compania_id.centro_trabajo_ids:
                datos += str(centro.codigo) + '|' + str(centro.nombre) + '|' + str(centro.direccion) + '|' + str(centro.zona) + '|' + str(centro.telefono) + '|' + str(centro.fax) + '|' + str(centro.nombre_contacto) + '|' + str(centro.correo_electronico) + '|' + str(centro.codigo_departamento) + '|' + str(centro.codigo_municipio) + '|' + str(centro.codigo_actividad_economica) + '\r\n'
            datos += '[tiposplanilla]' + '\r\n'
            for tipo_planilla in compania_id.tipo_planilla_ids:
                datos += str(tipo_planilla.codigo) + '|' + str(tipo_planilla.name) + '|' + str(tipo_planilla.tipo_afiliado) + '|' + str(tipo_planilla.periodo_planilla) + '|' + str(tipo_planilla.codigo_departamento) + '|' + str(tipo_planilla.codigo_actividad_economica) + '|' + str(tipo_planilla.clase_planilla) + '|' + str(tipo_planilla.tiempo_contrato) + '|' + '\r\n'
            datos += '[liquidaciones]' + '\r\n'
            for liquidacion in compania_id.tipo_planilla_ids.liquidaciones_ids:
                fecha_inicial = format_date(self.env, self.fecha_inicial, date_format="dd/MM/yyyy")
                fecha_final = format_date(self.env, self.fecha_final, date_format="dd/MM/yyyy")
                datos += str(liquidacion.numero) + '|' + str(liquidacion.tipo_planilla_id.codigo) + '|' + str(fecha_inicial) + '|' + str(fecha_final) + '|' + str(liquidacion.complementaria_original) + '|' + str(liquidacion.numero_nota_cargo) + '\r\n'
            datos += '[empleados]' + '\r\n'
            empleados = {}
            suspensiones = []
            for payslip_run in w.payslip_run_id:
                for slip in payslip_run.slip_ids:
                    if slip.employee_id.id not in empleados:
                        empleados[slip.employee_id.id] = {'empleado_id': slip.employee_id.id,'informacion': [0] * 19, 'suspension': ''}

                    numero_liquidacion = slip.employee_id.numero_liquidacion if slip.employee_id.numero_liquidacion else ''
                    numero_afiliado = slip.employee_id.igss if slip.employee_id.igss else ''
                    primer_nombre = slip.employee_id.primer_nombre if slip.employee_id.primer_nombre else ''
                    segundo_nombre = slip.employee_id.segundo_nombre if slip.employee_id.segundo_nombre else ''
                    primer_apellido = slip.employee_id.primer_apellido if slip.employee_id.primer_apellido else ''
                    segundo_apellido = slip.employee_id.segundo_apellido if slip.employee_id.segundo_apellido else ''
                    apellido_casada = slip.employee_id.apellido_casada if slip.employee_id.apellido_casada else ''
                    tipo_salario = slip.employee_id.tipo_salario if slip.employee_id.tipo_salario else ''
                    tiempo_contrato = slip.employee_id.tiempo_contrato if slip.employee_id.tiempo_contrato else ''

                    horas_laboradas = ''
                    dias_laborados = 0
                    for linea in slip.worked_days_line_ids:
                        if linea.work_entry_type_id in slip.employee_id.company_id.igss_dias_trabajo:
                            dias_laborados = linea.number_of_days

                    sueldo = 0
                    for linea in slip.line_ids:
                        if linea.salary_rule_id.id in slip.employee_id.company_id.sueldo_igss_ids.ids:
                            sueldo += linea.total

                    mes_inicio_contrato = employee_id.date_start.month
                    anio_inicio_contrato = employee_id.date_start.year
                    mes_final_contrato = employee_id.date_end.month if employee_id.date_end else ''
                    anio_final_contrato = employee_id.date_end.year if employee_id.date_end else ''
                    mes_planilla = payslip_run.date_start.month
                    anio_planilla = payslip_run.date_start.year
                    fecha_alta = format_date(self.env, employee_id.date_start, date_format="dd/MM/yyyy") if (mes_inicio_contrato == mes_planilla and anio_inicio_contrato == anio_planilla) else ''
                    fecha_baja = format_date(self.env, employee_id.date_end, date_format="dd/MM/yyyy") if (mes_final_contrato == mes_planilla and anio_final_contrato == anio_planilla) else ''

                    centro_trabajo = slip.employee_id.codigo_centro_trabajo if slip.employee_id.codigo_centro_trabajo else ''
                    nit = slip.employee_id.work_contact_id.vat if slip.employee_id.work_contact_id.vat else slip.employee_id.nit or ''
                    codigo_ocupacion = slip.employee_id.codigo_ocupacion if slip.employee_id.codigo_ocupacion else ''
                    condicion_laboral = slip.employee_id.condicion_laboral if slip.employee_id.condicion_laboral else ''
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

                    ausencias = self.env['hr.leave'].search([('employee_id','=', empleado['empleado_id']), ('request_date_from','>=',lote_id.fecha_inicial), ('request_date_to','<=',lote_id.fecha_final), ('state','=','validate')])
                    if ausencias:
                        for ausencia in ausencias:
                            if ausencia.holiday_status_id == self.env.ref('rrhh.suspension_igss') or ausencia.holiday_status_id.suspension_igss:
                                fecha_inicio = str(format_date(self.env, ausencia.request_date_from, date_format="dd/MM/yyyy"))
                                fecha_fin = str(format_date(self.env, ausencia.request_date_to, date_format="dd/MM/yyyy"))
                                igss = ausencia.employee_id.igss if ausencia.employee_id.igss else ""
                                primer_nombre = ausencia.employee_id.primer_nombre if ausencia.employee_id.primer_nombre else ""
                                segundo_nombre = ausencia.employee_id.segundo_nombre if ausencia.employee_id.segundo_nombre else ""
                                primer_apellido = ausencia.employee_id.primer_apellido if ausencia.employee_id.primer_apellido else ""
                                segundo_apellido = ausencia.employee_id.segundo_apellido if ausencia.employee_id.segundo_apellido else ""
                                apellido_casada = ausencia.employee_id.apellido_casada if ausencia.employee_id.apellido_casada else ""
                                suspensiones.append(numero_liquidacion + '|' + igss + '|' + primer_nombre + '|' + segundo_nombre + '|' + primer_apellido + '|' + segundo_apellido + '|' + apellido_casada  + '|' + str(fecha_inicio) + '|' + str(fecha_fin) + '|' + '\r\n')

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
            'res_model': 'rrhh.archivo_igss.wizard',
            'res_id': self.id,
            'view_id': False,
            'type': 'ir.actions.act_window',
            'target': 'new',
        }
