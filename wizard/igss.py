# -*- encoding: utf-8 -*-

from openerp import models, fields, api, _
from openerp.exceptions import UserError, ValidationError
import time
import base64
import xlwt
import io
import logging
import datetime
from datetime import datetime

class rrhh_igss_wizard(models.TransientModel):
    _name = 'rrhh.igss.wizard'

    def _default_payslip_run(self):
        logging.warn(self.env.context.get('active_ids'))
        if len(self.env.context.get('active_ids', [])) > 0:
            return self.env.context.get('active_ids')[0]
        else:
            return None

    payslip_run_id = fields.Many2one('hr.payslip.run', string='Payslip run',default=_default_payslip_run)
    archivo = fields.Binary('Archivo')
    name =  fields.Char('File Name', size=32)
    identificacion_tipo_planilla = fields.Char('Identificación tipo de planilla')
    nombre_tipo_planilla = fields.Char('Nombre tipo de planilla')
    tipo_afiliados = fields.Char('Tipo de afiliado')
    periodo_planilla = fields.Char('Periodo de planilla')
    departamento_republica = fields.Char('Departamento de la república donde laboran los empleados ')
    actividad_economica = fields.Char('Actividad económica')
    clase_planilla = fields.Char('Clase de planilla')
    numero_liquidacion = fields.Char('Numero de liquidacion')
    tipo_planilla_liquidacion = fields.Char('Tipo de planilla de liquidación')
    fecha_inicial = fields.Char('Fecha inicial liquidación')
    fecha_final = fields.Char('Fecha final de liquidación')
    tipo_liquidacion = fields.Char('Tipo de liquidación')
    numero_nota_cargo = fields.Char('Número nota de cargo')

    def generar(self):
        datos = ''
        for w in self:
            datos += str(w.payslip_run_id.slip_ids[0].company_id.version_mensaje) + '|' + str(datetime.today().strftime('%d/%m/%Y')) + '|' + str(w.payslip_run_id.slip_ids[0].company_id.numero_patronal) + '|'+ str(datetime.strptime(str(w.payslip_run_id.date_start),'%Y-%m-%d').date().strftime('%m')).lstrip('0')+ '|' + str(datetime.strptime(str(w.payslip_run_id.date_start),'%Y-%m-%d').date().strftime('%Y')).lstrip('0') + '|' + str(w.payslip_run_id.slip_ids[0].company_id.name) + '|' +str(w.payslip_run_id.slip_ids[0].company_id.vat) + '|'+ str(w.payslip_run_id.slip_ids[0].company_id.email) + '|' + str(w.payslip_run_id.slip_ids[0].company_id.tipo_planilla) + '\r\n'
            datos += '[centros]' + '\r\n'
            for centro in w.payslip_run_id.slip_ids[0].company_id.centro_trabajo_ids:
                datos += str(centro.codigo) + '|' + str(centro.nombre) + '|' + str(centro.direccion) + '|' + str(centro.zona) + '|' + str(centro.telefono) + '|' + str(centro.fax) + '|' + str(centro.nombre_contacto) + '|' + str(centro.correo_electronico) + '|' + str(centro.codigo_departamento) + '|' + str(centro.codigo_municipio) + '|' + str(centro.codigo_actividad_economica) + '\r\n'
            datos += '[tiposplanilla]' + '\r\n'
            # datos += str(w.payslip_run_id.slip_ids[0].company_id.identificacion_tipo_planilla) + '|' + str(w.payslip_run_id.slip_ids[0].company_id.nombre_tipo_planilla) + '|' + str(w.payslip_run_id.slip_ids[0].company_id.tipo_afiliados) + '|' + str(w.payslip_run_id.slip_ids[0].company_id.periodo_planilla) + '|' + str(w.payslip_run_id.slip_ids[0].company_id.departamento_republica) + '|' + str(w.payslip_run_id.slip_ids[0].company_id.actividad_economica) + '|' + str(w.payslip_run_id.slip_ids[0].company_id.clase_planilla) + '\r\n'
            datos += self.identificacion_tipo_planilla + '|' + self.nombre_tipo_planilla + '|' + self.tipo_afiliados + '|' + self.periodo_planilla + '|' + self.departamento_republica + '|' + self.actividad_economica + '|' + self.clase_planilla + '|' +'\r\n'
            datos += '[liquidaciones]' + '\r\n'
            datos += self.numero_liquidacion + '|' + self.tipo_planilla_liquidacion + '|' + self.fecha_inicial + '|' + self.fecha_final + '|' + self.tipo_liquidacion + '|' + (self.numero_nota_cargo if self.numero_nota_cargo else '') + '|' +'\r\n'
            datos += '[empleados]' + '\r\n'
            for slip in w.payslip_run_id.slip_ids:
                fecha_planilla = datetime.strptime(str(w.payslip_run_id.date_start), '%Y-%m-%d')
                mes_planilla = fecha_planilla.month
                anio_planilla = fecha_planilla.year
                contrato_ids = self.env['hr.contract'].search( [['employee_id', '=', slip.employee_id.id]],offset=0,limit=1,order='date_start desc')
                logging.warn(contrato_ids)
                datos += '1' + '|' + str(slip.employee_id.igss) + '|' + slip.employee_id.primer_nombre + '|'+(slip.employee_id.segundo_nombre if slip.employee_id.segundo_nombre else '' ) + '|' + (slip.employee_id.primer_apellido if slip.employee_id.primer_apellido else '' )+ '|' + (slip.employee_id.segundo_apellido if slip.employee_id.segundo_apellido else '') + '|'+ (slip.employee_id.apellido_casada if slip.employee_id.apellido_casada else '') + '|'
                if contrato_ids:
                    contrato = self.env['hr.contract'].browse([contrato_ids.id])
                    if contrato.date_end:
                        mes_contrato= datetime.strptime(contrato.date_end, '%Y-%m-%d')
                        mes_final_contrato = mes_contrato.month
                        anio_final_contrato = mes_contrato.year
                        if mes_planilla == mes_final_contrato and anio_final_contrato == anio_planilla:
                            datos += str(contrato.wage) + '|' + str(datetime.strptime(contrato.date_start,'%Y-%m-%d').date().strftime('%d/%m/%Y')) + '|' + str(datetime.strptime(contrato.date_end,'%Y-%m-%d').date().strftime('%d/%m/%Y')) + '|'
                    else:
                        mes_contrato = datetime.strptime(str(contrato.date_start), '%Y-%m-%d')
                        mes_final_contrato = mes_contrato.month
                        anio_final_contrato = mes_contrato.year
                        if mes_final_contrato == mes_planilla and anio_final_contrato == anio_planilla:
                            datos += str(contrato.wage) + '|' + str(datetime.strptime(contrato.date_start,'%Y-%m-%d').date().strftime('%d/%m/%Y')) + '|' + '|'
                        else:
                            datos += str(contrato.wage) + '|' + '|' + '' + '|'
                else:
                    datos += '|' + '|' + '|'
                datos += str(slip.employee_id.centro_trabajo_id.codigo) + '|' + str(slip.employee_id.nit) + '|' + str(slip.employee_id.codigo_ocupacion) + '|' + str(slip.employee_id.condicion_laboral) + '|' + '|' + '\r\n'
            datos += '[suspendidos]' + '\r\n'
            datos += '[licencias]' + '\r\n'
            datos += '[juramento]' + '\r\n'
            datos += 'BAJO MI EXCLUSIVA Y ABSOLUTA RESPONSABILIDAD, DECLARO QUE LA INFORMACION QUE AQUI CONSIGNO ES FIEL Y EXACTA, QUE ESTA PLANILLA INCLUYE A TODOS LOS TRABAJADORES QUE ESTUVIERON A MI SERVICIO Y QUE SUS SALARIOS SON LOS EFECTIVAMENTE DEVENGADOS, DURANTE EL MES ARRIBA INDICADO' + '\r\n'
            datos += '[finplanilla]' + '\r\n'
            datos = datos.replace('False', '')
        datos = base64.b64encode(datos.encode("utf-8"))
        self.write({'archivo': datos, 'name':'igss.txt'})

        return {
            'view_type': 'form',
            'view_mode': 'form',
            'res_model': 'rrhh.igss.wizard',
            'res_id': self.id,
            'view_id': False,
            'type': 'ir.actions.act_window',
            'target': 'new',
        }
