# -*- coding: utf-8 -*-
# Part of Odoo. See LICENSE file for full copyright and licensing details.

from odoo import api, fields, models, _
from odoo.exceptions import UserError
import xlsxwriter
import base64
import io
import logging
from datetime import datetime, date, time

class rrhh_informe_empleador(models.TransientModel):
    _name = 'rrhh.informe_empleador.wizard'
    _description = 'Wizard para generar informe de empleador'

    anio = fields.Integer('Año', required=True)
    name = fields.Char('Nombre archivo')
    archivo = fields.Binary('Archivo')

    def empleados_inicio_anio(self, company_id, anio):
        empleados = 0
        for empleado in self.env['hr.employee'].search([('company_id', '=', company_id)]):
            anio_fin = 0
            anio_inicio = empleado.date_start.year
            if empleado.date_end:
                anio_fin = empleado.date_end.year
            if anio_inicio < anio and (empleado.date_end == False or anio_fin < anio) :
                empleados += 1

        return empleados

    def empleados_fin_anio(self, company_id, anio):
        empleados = 0
        for empleado in self.env['hr.employee'].search([['company_id', '=', company_id]]):
            anio_fin = 0
            anio_inicio = empleado.date_start.year
            if empleado.date_end:
                anio_fin = empleado.date_end.year
            if anio_inicio <= anio and (empleado.date_end == False or anio_fin <= anio) :
                empleados += 1

        return empleados

    def calcular_indemnizacion(self, empleado_id, anio):
        dias_laborados = 0
        salario_promedio = 0
        indemnizacion = 0
        regla_76_78 = 0
        regla_42_92 = 0
        indemnizacion = 0
        anio_fin = date(anio, 12, 31)
        if empleado_id.calcula_indemnizacion:
            dias_laborados = self.env['hr.payslip'].dias_trabajados_rango(empleado_id.date_start, empleado_id.date_end or anio_fin)
            salario_promedio = self.env['hr.payslip'].salario_promedio(empleado_id, empleado_id.date_end or anio_fin)
            salario_diario = salario_promedio / 365
            regla_76_78 = ((salario_promedio / 12) / 365) * dias_laborados
            regla_42_92 = ((salario_promedio / 12) / 365) * dias_laborados
            indemnizacion = (salario_diario * dias_laborados) + regla_76_78 + regla_42_92

        return indemnizacion

    def dias_trabajados_anual(self, empleado_id, anio):
        anio_inicio_contrato = empleado_id.date_start.year

        fecha_inicio = max(date(anio, 1, 1), empleado_id.date_start) if empleado_id.date_start else date(anio, 1, 1)
        fecha_fin = min(date(anio, 12, 31), empleado_id.date_end) if empleado_id.date_end else date(anio, 12, 31)

        dias = empleado_id._get_work_days_data_batch(datetime.combine(fecha_inicio, datetime.min.time()), datetime.combine(fecha_fin, datetime.max.time()), calendar=empleado_id.resource_calendar_id)
        return dias[empleado_id.id]['days']

    def print_report_excel(self):
        for w in self:
            empleados_id = self.env.context.get('active_ids', [])

            f = io.BytesIO()
            libro = xlsxwriter.Workbook(f)
            formato_fecha = libro.add_format({'num_format': 'dd/mm/yy'})
            empleados_archivados = self.env['hr.employee'].sudo().search([('active','=',False),('id', 'in', empleados_id)])
            empleados_activos = self.env['hr.employee'].sudo().search([('active','=',True),('id', 'in', empleados_id)])
            empleados = empleados_archivados + empleados_activos
            responsable_id = self.env['hr.employee'].sudo().search([['id', '=', self.env.user.id]])
            datos_compania = self.env.company

            hoja_patrono = libro.add_worksheet('Patrono')
            empleados_inicio_anio = self.empleados_inicio_anio(datos_compania.id, w['anio'])
            empleados_fin_anio = self.empleados_fin_anio(datos_compania.id, w['anio'])

            hoja_patrono.write(6, 0, 'Datos De Identificación')
            hoja_patrono.write(7, 0, 'Nit')
            hoja_patrono.write(7, 1, datos_compania.vat)
            hoja_patrono.write(8, 0, 'Nombre de la empresa')
            hoja_patrono.write(8, 1, datos_compania.company_registry)
            hoja_patrono.write(9, 0, 'Nacionalidad del empleador')
            hoja_patrono.write(9, 1, datos_compania.country_id.name)
            hoja_patrono.write(10, 0, 'Denominación o razón social de patrono')
            hoja_patrono.write(10, 1, datos_compania.name)
            hoja_patrono.write(11, 0, 'Numero patronal IGSS')
            hoja_patrono.write(11, 1, datos_compania.numero_patronal)

            hoja_patrono.write(12, 0, 'Datos General')
            hoja_patrono.write(13, 0, 'Barrio o Colonia')
            hoja_patrono.write(13, 1, datos_compania.barrio_colonia)
            hoja_patrono.write(13, 2, 'Zona')
            hoja_patrono.write(13, 3, datos_compania.zona)
            hoja_patrono.write(14, 0, 'Calle')
            hoja_patrono.write(14, 1, datos_compania.street2)
            hoja_patrono.write(14, 2, 'Avenida')
            hoja_patrono.write(14, 3, datos_compania.street)
            hoja_patrono.write(15, 0, 'Teléfono')
            hoja_patrono.write(15, 1, datos_compania.phone)
            hoja_patrono.write(15, 2, 'Nomenclatura')
            hoja_patrono.write(15, 3, datos_compania.nomenclatura)
            hoja_patrono.write(16, 0, 'Sitio Web')
            hoja_patrono.write(16, 1, datos_compania.website)
            hoja_patrono.write(16, 2, 'E-Mail')
            hoja_patrono.write(16, 3, datos_compania.email)
            hoja_patrono.write(17, 0, 'Existe Sindicato (SI) O (NO)')
            hoja_patrono.write(17, 1, datos_compania.sindicato)

            hoja_patrono.write(19, 0, 'Ubicación Geográfica')
            hoja_patrono.write(20, 0, 'País')
            hoja_patrono.write(20, 1, datos_compania.country_id.name)
            hoja_patrono.write(20, 2, 'Región')
            hoja_patrono.write(20, 3, datos_compania.state_id.name)
            hoja_patrono.write(21, 0, 'Departamento')
            hoja_patrono.write(21, 1, datos_compania.state_id.name)
            hoja_patrono.write(21, 2, 'Municipio')
            hoja_patrono.write(21, 3, datos_compania.city)
            hoja_patrono.write(22, 0, 'Datos Económicos')
            hoja_patrono.write(23, 0, 'Año de Inicio de Operaciones')
            hoja_patrono.write(23, 1, datos_compania.anio_inicio_operaciones)
            hoja_patrono.write(24, 0, 'Cantidad Total de Empleados Inicio de Año ')
            hoja_patrono.write(24, 1, empleados_inicio_anio)
            hoja_patrono.write(25, 0, 'Cantidad Total de Empleados fin de Año')
            hoja_patrono.write(25, 1, empleados_fin_anio)
            hoja_patrono.write(26, 0, 'Tamaño de la empresa por ventas anuales en salarios minimos')
            hoja_patrono.write(26, 1, datos_compania.tamanio_empresa_ventas)
            hoja_patrono.write(27, 0, 'Tamaño de empresa según cantidad de Trabajadores')
            hoja_patrono.write(27, 1, datos_compania.tamanio_empresa_trabajadores)
            hoja_patrono.write(28, 0, 'Tiene planificado contratar nuevo personal (SI) (NO)')
            hoja_patrono.write(28, 1, datos_compania.contratar_personal)
            hoja_patrono.write(29, 0, 'Contabilidad Completa')
            hoja_patrono.write(29, 1, datos_compania.contabilidad_completa)

            hoja_patrono.write(31, 0, 'Actividad Económica Principal')
            hoja_patrono.write(32, 0, 'Actividad Gran Grupo')
            hoja_patrono.write(32, 1, datos_compania.actividad_gran_grupo)
            hoja_patrono.write(33, 0, 'Actividad Económica')
            hoja_patrono.write(33, 1, datos_compania.actividad_economica)
            hoja_patrono.write(34, 0, 'Sub Actividad Económica')
            hoja_patrono.write(34, 1, datos_compania.sub_actividad_economica)
            hoja_patrono.write(35, 0, 'Ocupación Grupo')
            hoja_patrono.write(35, 1, datos_compania.ocupacion_grupo)

            hoja_patrono.write(37, 0, 'Datos Del Contacto')
            hoja_patrono.write(38, 0, 'Nombre Del Represéntate. Legal')
            hoja_patrono.write(38, 1, datos_compania.representante_legal_id.name)
            hoja_patrono.write(39, 0, 'Tipo De Documento Del Represéntate. Legal ')
            hoja_patrono.write(39, 1, 'DPI')
            hoja_patrono.write(40, 0, 'Nombre Jefe De Recursos Humanos')
            hoja_patrono.write(40, 1, datos_compania.jefe_recursos_humanos_id.name)
            hoja_patrono.write(41, 0, 'No. De Identificación De  Jefe De RR.HH.')
            hoja_patrono.write(41, 1, datos_compania.jefe_recursos_humanos_id.identification_id)
            hoja_patrono.write(42, 0, 'E-Mail Del Jefe RR.HH.')
            hoja_patrono.write(42, 1, datos_compania.jefe_recursos_humanos_id.work_email)
            hoja_patrono.write(43, 0, 'E-Mail Del Responsable Del Informe ')
            hoja_patrono.write(43, 1, responsable_id.work_email)
            hoja_patrono.write(44, 0, 'Teléfono Del Represéntate Del Informe')
            hoja_patrono.write(44, 1, responsable_id.work_phone)
            hoja_patrono.write(45, 0, 'Nacionalidad Del Representante Legal')
            hoja_patrono.write(45, 1, datos_compania.representante_legal_id.country_id.name)
            hoja_patrono.write(46, 0, 'No. De Identificación Del Represéntate Legal')
            hoja_patrono.write(46, 1, datos_compania.representante_legal_id.identification_id)
            hoja_patrono.write(47, 0, 'Tipo De Documentación Del Jefe De RR.HH.')
            hoja_patrono.write(47, 1, 'DPI')
            hoja_patrono.write(48, 0, 'Teléfono Jefe RR.HH.')
            hoja_patrono.write(48, 1, datos_compania.jefe_recursos_humanos_id.work_phone)
            hoja_patrono.write(49, 0, 'Nombre Del Represéntate de Elaborar el Informe Del Empleador')
            hoja_patrono.write(49, 1, responsable_id.name)
            hoja_patrono.write(50, 0, 'Documento Identificación Responsable')
            hoja_patrono.write(50, 1, responsable_id.identification_id)
            hoja_patrono.write(51, 0, 'Nacionalidad Del Responsable')
            hoja_patrono.write(51, 1, responsable_id.country_id.name)
            hoja_patrono.write(52, 0, 'Año Del Informe ')
            hoja_patrono.write(52, 1, w['anio'])

            hoja_empleado = libro.add_worksheet('Empleado')
            hoja_empleado.write(0, 0, 'Numero de empleado')
            hoja_empleado.write(0, 1, 'Primer Nombre')
            hoja_empleado.write(0, 2, 'Segundo Nombre')
            hoja_empleado.write(0, 3, 'Tercer Nombre')
            hoja_empleado.write(0, 4, 'Primer Apellido')
            hoja_empleado.write(0, 5, 'Segundo Apellido')
            hoja_empleado.write(0, 6, 'Apellido de casada')
            hoja_empleado.write(0, 7, 'Nacionalidad')
            hoja_empleado.write(0, 8, 'Tipo de discapacidad')
            hoja_empleado.write(0, 9, 'Estado Civil')
            hoja_empleado.write(0, 10, 'Documento identificación (DPI, Pasaporte u otro)')
            hoja_empleado.write(0, 11, 'Número de Documento')
            hoja_empleado.write(0, 12, 'Pais Origen')
            hoja_empleado.write(0, 13, 'Número de expediente del permiso de extranjero')
            hoja_empleado.write(0, 14, 'Lugar Nacimiento')
            hoja_empleado.write(0, 15, 'Número de Identificación Tributaria NIT')
            hoja_empleado.write(0, 16, 'Número de Afiliación IGSS')
            hoja_empleado.write(0, 17, 'Sexo (M) O (F)')
            hoja_empleado.write(0, 18, 'Fecha Nacimiento')
            hoja_empleado.write(0, 19, 'Nivel Academico')
            hoja_empleado.write(0, 20, 'Titulo o diploma (profesión)')
            hoja_empleado.write(0, 21, 'Pueblo de pertenencia')
            hoja_empleado.write(0, 22, 'Comunidad')
            hoja_empleado.write(0, 23, 'Cantidad de Hijos')
            hoja_empleado.write(0, 24, 'Temporalidad del contrato')
            hoja_empleado.write(0, 25, 'Tipo de contrato')
            hoja_empleado.write(0, 26, 'Fecha Inicio Labores')
            hoja_empleado.write(0, 27, 'Fecha de reinicio de labores')
            hoja_empleado.write(0, 28, 'Fecha de finalización de labores')
            hoja_empleado.write(0, 29, 'Ocupación')
            hoja_empleado.write(0, 30, 'Jornada de Trabajo')
            hoja_empleado.write(0, 31, 'Dias Laborados en el Año')
            hoja_empleado.write(0, 32, 'Salario Mensual Nominal')
            hoja_empleado.write(0, 33, 'Salario Anual Nominal')
            hoja_empleado.write(0, 34, 'Bonificación Decreto 78-89  (Q.250.00)')
            hoja_empleado.write(0, 35, 'Total Horas Extras Anuales')
            hoja_empleado.write(0, 36, 'Valor de Hora Extra')
            hoja_empleado.write(0, 37, 'Monto Aguinaldo Decreto 76-78')
            hoja_empleado.write(0, 38, 'Monto Bono 14  Decreto 42-92')
            hoja_empleado.write(0, 39, 'Retribución por Comisiones')
            hoja_empleado.write(0, 40, 'Viaticos')
            hoja_empleado.write(0, 41, 'Bonificaciones Adicionales')
            hoja_empleado.write(0, 42, 'Retribución por vacaciones')
            hoja_empleado.write(0, 43, 'Retribución por Indemnización (Articulo 82)')
            hoja_empleado.write(0, 44, 'Sucursal')

            fila = 1
            empleado_numero = 1
            for empleado in empleados:
                if empleado.primer_nombre:
                    nominas_lista = []
                    nomina_id = self.env['hr.payslip'].search([['employee_id', '=', empleado.id]])
                    dias_trabajados = 0
                    salario_anual_nominal = 0
                    bonificacion = 0
                    estado_civil = 0
                    horas_extras = 0
                    aguinaldo = 0
                    bono = 0
                    bonificaciones_adicionales = 0
                    valor_horas_extras = 0
                    retribucion_comisiones = 0
                    viaticos = 0
                    retribucion_vacaciones = 0
                    bonificacion_decreto = 0
                    precision_currency = empleado.company_id.currency_id
                    indemnizacion = precision_currency.round(self.calcular_indemnizacion(empleado.id, w['anio'])) if empleado.date_end else 0
                    salario_anual_nominal_promedio = 0
                    nominas = {}
                    numero_horas_extra = 0
                    numero_nominas_salario = 0
                    genero = ''
                    for nomina in nomina_id:
                        nomina_anio = nomina.date_to.year
                        nomina_mes = nomina.date_to.month
                        if w['anio'] == nomina_anio:
                            if nomina.input_line_ids:
                                for entrada in nomina.input_line_ids:
                                    for horas_entrada in nomina.company_id.numero_horas_extras_ids:
                                        if entrada.code == horas_entrada.code:
                                            numero_horas_extra += entrada.amount
                            for linea in nomina.worked_days_line_ids:
                                if linea.work_entry_type_id.code == empleado.company_id.tipo_entrada_trabajo_id.code:
                                    dias_trabajados += linea.number_of_days
                            for linea in nomina.line_ids:
                                if linea.salary_rule_id.id in nomina.company_id.salario_ids.ids:
                                    salario_anual_nominal += linea.total
                                    if nomina_mes not in nominas:
                                        nominas[nomina_mes] = {'salario': 0,'bonificacion':0}
                                    nominas[nomina_mes]['salario'] += salario_anual_nominal
                                    numero_nominas_salario += 1
                                if linea.salary_rule_id.id in nomina.company_id.bonificacion_ids.ids:
                                    bonificacion += linea.total
                                if linea.salary_rule_id.id in nomina.company_id.aguinaldo_ids.ids:
                                    aguinaldo += linea.total
                                if linea.salary_rule_id.id in nomina.company_id.bono_ids.ids:
                                    bono += linea.total
                                if linea.salary_rule_id.id in nomina.company_id.horas_extras_ids.ids:
                                    horas_extras += linea.total
                                if linea.salary_rule_id.id in nomina.company_id.retribucion_comisiones_ids.ids:
                                    retribucion_comisiones += linea.total
                                if linea.salary_rule_id.id in nomina.company_id.viaticos_ids.ids:
                                    viaticos += linea.total
                                if linea.salary_rule_id.id in nomina.company_id.retribucion_vacaciones_ids.ids:
                                    retribucion_vacaciones += linea.total
                                if linea.salary_rule_id.id in nomina.company_id.bonificaciones_adicionales_ids.ids:
                                    bonificaciones_adicionales += linea.total
                                if linea.salary_rule_id.id in nomina.company_id.decreto_ids.ids:
                                    bonificacion_decreto += linea.total
                                    if nomina_mes not in nominas:
                                        nominas[nomina_mes] = {'salario': 0,'bonificacion':0}
                                    nominas[nomina_mes]['bonificacion'] += bonificacion_decreto

                    salario_anual_nominal_promedio = salario_anual_nominal / len(nominas) if salario_anual_nominal > 0 else 0
                    if empleado.sex == 'male':
                        genero = '1'
                    if empleado.sex == 'female':
                        genero = '2'
                    if empleado.marital == 'single':
                        estado_civil = 1
                    if empleado.marital == 'married':
                        estado_civil = 2
                    if empleado.marital == 'widower':
                        estado_civil = 3
                    if empleado.marital == 'divorced':
                        estado_civil = 4
                    if empleado.marital == 'separado':
                        estado_civil = 5
                    if empleado.marital == 'cohabitant':
                        estado_civil = 6
                    dias_trabajados_anual = self.dias_trabajados_anual(empleado, w['anio'])
                    hoja_empleado.write(fila, 0, empleado_numero)
                    hoja_empleado.write(fila, 1, empleado.primer_nombre or '')
                    hoja_empleado.write(fila, 2, empleado.segundo_nombre or '')
                    hoja_empleado.write(fila, 3, empleado.tercer_nombre or '')
                    hoja_empleado.write(fila, 4, empleado.primer_apellido or '')
                    hoja_empleado.write(fila, 5, empleado.segundo_apellido or '')
                    hoja_empleado.write(fila, 6, empleado.apellido_casada or '')
                    hoja_empleado.write(fila, 7, empleado.nacionalidad or '')
                    hoja_empleado.write(fila, 8, empleado.tipo_discapacidad or '')
                    hoja_empleado.write(fila, 9, estado_civil)
                    hoja_empleado.write(fila, 10, empleado.documento_identificacion or '')
                    hoja_empleado.write(fila, 11, empleado.identification_id or '')
                    hoja_empleado.write(fila, 12, empleado.country_of_birth.code or '')
                    hoja_empleado.write(fila, 13, empleado.permiso_trabajo or '')
                    hoja_empleado.write(fila, 14, empleado.codigo_municipio_nacimiento or '')
                    hoja_empleado.write(fila, 15, empleado.work_contact_id.vat or empleado.nit or '')
                    hoja_empleado.write(fila, 16, empleado.igss or '')
                    hoja_empleado.write(fila, 17, genero)
                    hoja_empleado.write(fila, 18, empleado.birthday or '', formato_fecha)
                    hoja_empleado.write(fila, 19, empleado.nivel_academico or '')
                    hoja_empleado.write(fila, 20, empleado.profesion or '')
                    hoja_empleado.write(fila, 21, empleado.pueblo_pertenencia or '')
                    hoja_empleado.write(fila, 22, empleado.comunidad_linguistica or '')
                    hoja_empleado.write(fila, 23, empleado.children or '')
                    hoja_empleado.write(fila, 24, empleado.temporalidad_contrato or '')
                    hoja_empleado.write(fila, 25, empleado.tipo_contrato or '')
                    hoja_empleado.write(fila, 26, empleado.date_start or '', formato_fecha)
                    hoja_empleado.write(fila, 27, empleado.fecha_reinicio_labores or '', formato_fecha)
                    hoja_empleado.write(fila, 28, empleado.date_end or '', formato_fecha)
                    hoja_empleado.write(fila, 29, empleado.codigo_ocupacion or '')
                    hoja_empleado.write(fila, 30, empleado.jornada_trabajo or '')
                    hoja_empleado.write(fila, 31, dias_trabajados_anual)
                    hoja_empleado.write(fila, 32, salario_anual_nominal_promedio)
                    hoja_empleado.write(fila, 33, salario_anual_nominal)
                    hoja_empleado.write(fila, 34, bonificacion_decreto)
                    hoja_empleado.write(fila, 35, numero_horas_extra)
                    hoja_empleado.write(fila, 36, ((horas_extras / numero_horas_extra) if numero_horas_extra > 0 else horas_extras))
                    hoja_empleado.write(fila, 37, aguinaldo)
                    hoja_empleado.write(fila, 38, bono)
                    hoja_empleado.write(fila, 39, retribucion_comisiones)
                    hoja_empleado.write(fila, 40, viaticos)
                    hoja_empleado.write(fila, 41, bonificaciones_adicionales)
                    hoja_empleado.write(fila, 42, retribucion_vacaciones)
                    hoja_empleado.write(fila, 43, indemnizacion)                    
                    hoja_empleado.write(fila, 44, empleado.sucursal or '')
                    
                    empleado_numero +=1

                    fila += 1

            libro.close()
            datos = base64.b64encode(f.getvalue())
            self.write({'archivo': datos, 'name': 'informe_del_empleador.xlsx'})

        return {
            'view_type': 'form',
            'view_mode': 'form',
            'res_model': 'rrhh.informe_empleador.wizard',
            'res_id': self.id,
            'view_id': False,
            'type': 'ir.actions.act_window',
            'target': 'new',
        }
