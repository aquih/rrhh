# -*- coding: utf-8 -*-
# Part of Odoo. See LICENSE file for full copyright and licensing details.

from odoo import api, fields, models, _
from odoo.exceptions import UserError
import base64
import xlsxwriter
import logging
import io
import datetime

class rrhh_informe_isr(models.TransientModel):
    _name = 'rrhh.informe_isr'

    tipo = fields.Selection([('proyeccion', 'Proyección'),
        ('actualizacion', 'Actualización'),
        ('ajuste_suspesion','Ajuste/Suspensión'),
        ('liquidacion_labor','Liquidacion fin labores'),
        ('liquidacion_periodo','Liquidación fin periodo'),
        ('retencion_pago','Retención por pago')],'Tipo', default="proyeccion")
    anio = fields.Integer('Año')
    fecha_inicio = fields.Date('Fecha inicio')
    fecha_fin = fields.Date('Fecha fin')
    archivo = fields.Binary('Archivo')
    name =  fields.Char('File Name', size=32)

    def _get_empleados(self, empleados):
        empleado_id = self.env['hr.employee'].search([('id', 'in', empleados)])
        return empleado_id

    def _get_informacion(self, empleado_id, fecha_inicio, fecha_fin, tipo):
        otros_ingresos = 0
        viaticos = 0
        igss_total = 0
        aguinaldo_anual = 0
        bono_anual = 0
        renta_patrono_actual = 0
        if tipo == "liquidacion_labor":
            nomina_id = self.env['hr.payslip'].search([('employee_id','=', empleado_id),('date_from', '>=', fecha_inicio),('date_to', '<=', fecha_fin)])
        elif tipo == "liquidacion_periodo":
            nomina_id = self.env['hr.payslip'].search([('employee_id','=', empleado_id),('date_from', '>=', fecha_inicio),('date_to', '<=', fecha_fin)])
        else:
            nomina_id = self.env['hr.payslip'].search([('employee_id','=', empleado_id),('date_to', '<', fecha_inicio)])
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

    def _get_retencion_pago(self, empleados, fecha_inicio, fecha_fin):
        nomina_id = self.env['hr.payslip'].search([('employee_id','in', empleados),('date_from', '>=', fecha_inicio),('date_to','<=',fecha_fin )])
        empleados_dic = {}
        for nomina in nomina_id:
            if nomina.line_ids:
                if nomina.employee_id.id not in empleados_dic:
                    empleados_dic[nomina.employee_id.id] = [0,0,0]
                for linea in nomina.line_ids:
                    if linea.salary_rule_id.id in nomina.employee_id.company_id.isr_ids.ids:
                        empleados_dic[nomina.employee_id.id][1] += linea.total
                    if linea.salary_rule_id.id in nomina.employee_id.company_id.base_gravada_ids.ids:
                        empleados_dic[nomina.employee_id.id][0] += linea.total
                    if linea.salary_rule_id.id in nomina.employee_id.company_id.ajuste_suspension_ids.ids:
                        empleados_dic[nomina.employee_id.id][2] += linea.total
        return empleados_dic

    def generar(self):
        datos = ''
        f = io.BytesIO()
        libro = xlsxwriter.Workbook(f)
        hoja = False
        hoja_carga_ajuste = False
        hoja_fin_labores = False
        hoja_fin_periodo = False
        hoja_retencion = False
        formato_fecha = libro.add_format({'num_format': 'dd/mm/yy'})
        if self.tipo == "proyeccion":
            hoja = libro.add_worksheet('Cargaproyecciones')
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
            hoja.write(0, 32, 'Retencion Ex Patrono 4proyeccion')
            hoja.write(0, 33, 'NIT ex patrono 5')
            hoja.write(0, 34, 'Renta Ex Patrono 5')
            hoja.write(0, 35, 'Retencion Ex Patrono 5')
            hoja.write(0, 36, 'Otros Ingresos Gravados')
            hoja.write(0, 37, 'Aguinaldo')
            hoja.write(0, 38, 'Bono Anual de Trabajadores')
            hoja.write(0, 39, 'Cuotas IGSS  y Otros Planes de Seguridad Social ')
            #hoja.write(0, 40, 'Fecha de actualización de Renta')
            fila = 1

            for empleado in self._get_empleados(self.env.context.get('active_ids', [])):
                hoja.write(fila, 0, empleado.nit if empleado.nit else '')
                hoja.write(fila, 1, empleado.name)

                if empleado.contract_id:
                    renta_patrono_actual = 0
                    bono_anual = 0
                    aguinaldo = 0
                    otros_ingresos_gravados = 0
                    cuota_igss = 0
                    fecha_alta = ""
                    anio_ingreso = empleado.contract_id.date_start.year
                    if anio_ingreso< self.anio:
                        renta_patrono_actual = ((empleado.contract_id.wage+250) * 12)
                    else:
                        dias_trabajados = (datetime.datetime.strptime(str(self.anio)+'-12-31', '%Y-%m-%d').date() - empleado.contract_id.date_start).days+1
                        valor_diario = (((empleado.contract_id.wage+250) * 12)/365)
                        renta_patrono_actual = valor_diario * dias_trabajados

                    fecha_bono_anterior = datetime.datetime.strptime(str(self.anio-1)+'-07-01', '%Y-%m-%d').date()
                    if empleado.contract_id.date_start < fecha_bono_anterior:
                        bono_anual = empleado.contract_id.wage
                    else:
                        dias_trabajados = (datetime.datetime.strptime(str(self.anio)+'-06-30', '%Y-%m-%d').date() - empleado.contract_id.date_start).days+1
                        valor_diario = (empleado.contract_id.wage /365)
                        bono_anual = valor_diario * dias_trabajados

                    fecha_inicio_calculo = datetime.datetime.strptime(str(self.anio-1)+'-12-01', '%Y-%m-%d').date()
                    if empleado.contract_id.date_start < fecha_inicio_calculo:
                        aguinaldo = empleado.contract_id.wage
                    else:
                        dias_trabajados = (datetime.datetime.strptime(str(self.anio)+'-11-30', '%Y-%m-%d').date() - empleado.contract_id.date_start).days+1
                        valor_diario = (empleado.contract_id.wage/365)
                        aguinaldo = valor_diario * dias_trabajados

                    if anio_ingreso< self.anio:
                        otros_ingresos_gravados = (empleado.contract_id.base_extra)
                    else:
                        dias_trabajados = (datetime.datetime.strptime(str(self.anio)+'-12-31', '%Y-%m-%d').date() - empleado.contract_id.date_start).days+1
                        valor_diario = ((empleado.contract_id.base_extra * 12)/365)
                        otros_ingresos_gravados = valor_diario * dias_trabajados

                    if anio_ingreso< self.anio:
                        # cuota_igss = (empleado.contract_id.wage)
                        fecha_alta = str(self.anio)+"-01-01"
                        fecha_alta = datetime.datetime.strptime(fecha_alta, '%Y-%m-%d').date()
                    else:
                        fecha_alta = empleado.contract_id.date_start
                    dias_trabajados = (datetime.datetime.strptime(str(self.anio)+'-12-31', '%Y-%m-%d').date() - fecha_alta).days+1
                    valor_diario = ((empleado.contract_id.wage * 12)/365)
                    cuota_igss = (valor_diario * dias_trabajados) * 0.0483
                    #logging.war
                    #fecha_alta = empleado.contract_id.date_start

                    # otra_info = self._get_informacion(empleado.id, '01-07-'+str(self.anio-1), '01-06-'+str(self.anio))
                    hoja.write(fila, 2, fecha_alta, formato_fecha)
                    hoja.write(fila, 3, renta_patrono_actual)
                    hoja.write(fila, 4, bono_anual)
                    hoja.write(fila, 5, aguinaldo)
                    hoja.write(fila, 36, otros_ingresos_gravados)
                    hoja.write(fila, 37, aguinaldo)
                    hoja.write(fila, 38, bono_anual)
                    hoja.write(fila, 39, cuota_igss)
                    #hoja.write(fila, 40, fecha_alta, formato_fecha)
                fila += 1

        if self.tipo == "actualizacion":
            hoja = libro.add_worksheet('Cargaactualizaciones')
            hoja.write(0, 0, 'NIT Empleado')
            hoja.write(0, 1, 'Nombre del Empleado')
            #hoja.write(0, 2, 'Fecha de Alta')
            hoja.write(0, 2, 'Renta Patrono Actual')
            hoja.write(0, 3, 'Bono Anual de Trabajadores')
            hoja.write(0, 4, 'Aguinaldo')
            hoja.write(0, 5, 'NIT Otro Patrono 1')
            hoja.write(0, 6, 'Renta Otro Patrono 1')
            hoja.write(0, 7, 'Retencion Otro Patrono 1')
            hoja.write(0, 8, 'NIT Otro Patrono 2')
            hoja.write(0, 9, 'Renta Otro Patrono 2')
            hoja.write(0, 10, 'Retencion Otro Patrono 2')
            hoja.write(0, 11, 'NIT Otro Patrono 3')
            hoja.write(0, 12, 'Renta Otro Patrono 3')
            hoja.write(0, 13, 'Retencion Otro Patrono 3')
            hoja.write(0, 14, 'NIT Otro Patrono 4')
            hoja.write(0, 15, 'Renta Otro Patrono 4')
            hoja.write(0, 16, 'Retencion Otro Patrono 4')
            hoja.write(0, 17, 'NIT Otro Patrono 5')
            hoja.write(0, 18, 'Renta Otro Patrono 5')
            hoja.write(0, 19, 'Retencion Otro Patrono 5')
            hoja.write(0, 20, 'NIT ex patrono 1')
            hoja.write(0, 21, 'Renta Ex Patrono 1')
            hoja.write(0, 22, 'Retencion Ex Patrono 1')
            hoja.write(0, 23, 'NIT ex patrono 2')
            hoja.write(0, 24, 'Renta Ex Patrono 2')
            hoja.write(0, 25, 'Retencion Ex Patrono 2')
            hoja.write(0, 26, 'NIT ex patrono 3')
            hoja.write(0, 27, 'Renta Ex Patrono 3')
            hoja.write(0, 28, 'Retencion Ex Patrono 3')
            hoja.write(0, 29, 'NIT ex patrono 4')
            hoja.write(0, 30, 'Renta Ex Patrono 4')
            hoja.write(0, 31, 'Retencion Ex Patrono 4')
            hoja.write(0, 32, 'NIT ex patrono 5')
            hoja.write(0, 33, 'Renta Ex Patrono 5')
            hoja.write(0, 34, 'Retencion Ex Patrono 5')
            hoja.write(0, 35, 'Otros Ingresos Gravados')
            hoja.write(0, 36, 'Aguinaldo')
            hoja.write(0, 37, 'Bono Anual de Trabajadores')
            hoja.write(0, 38, 'Cuotas IGSS  y Otros Planes de Seguridad Social ')
            hoja.write(0, 39, 'Fecha de actualización de Renta')
            fila = 1

            for empleado in self._get_empleados(self.env.context.get('active_ids', [])):
                if empleado.contract_id:
                    if len(empleado.contract_id.historial_salario_ids) > 1:
                        salario = 0
                        fecha_actualizacion = ""
                        posicion_salario = 0
                        salario_anterior = 0
                        for linea_historial in empleado.contract_id.historial_salario_ids:
                            if linea_historial.fecha and ( linea_historial.fecha >= self.fecha_inicio and linea_historial.fecha <= self.fecha_fin):
                                posicion_salario = empleado.contract_id.historial_salario_ids.ids.index(linea_historial.id)
                                salario = linea_historial.salario
                                fecha_actualizacion = linea_historial.fecha

                        salario_anterior = empleado.contract_id.historial_salario_ids[posicion_salario-1].salario
                        if salario > 0:
                            hoja.write(fila, 0, empleado.nit if empleado.nit else '')
                            hoja.write(fila, 1, empleado.name)
                            renta_patrono_actual = 0
                            bono_anual = 0
                            aguinaldo = 0
                            otros_ingresos_gravados = 0
                            cuota_igss = 0
                            anio_actual = self.fecha_fin.year
                            anio_ingreso = empleado.contract_id.date_start.year
                            empleado_planillas = self._get_informacion(empleado.id, self.fecha_inicio , '01/01/'+str(anio_actual), 'actualizacion')

                            fecha_final_date = datetime.datetime.strptime(str(anio_actual)+'-12-31', '%Y-%m-%d').date()
                            # diferencia_meses = (fecha_final_date.year - self.fecha_inicio.year) * 12 + (fecha_final_date.month - self.fecha_inicio.month)
                            dias_trabajados_proyectados = (fecha_final_date - self.fecha_inicio).days + 1
                            renta_patrono_actual = empleado_planillas["renta_patrono_actual"] +  ((((salario+250)*12)/365) * dias_trabajados_proyectados)

                            # ---------------------------- INICIO CALCULO BONO
                            fecha_final_nuevo_salario_bono = datetime.datetime.strptime(str(anio_actual)+'-06-30', '%Y-%m-%d').date()
                            dias_nuevo_salario_bono = (fecha_final_nuevo_salario_bono - self.fecha_inicio).days + 1
                            bono_nuevo_salario = ((salario * 12)/ 365)*dias_nuevo_salario_bono

                            fecha_inicio_antiguo_salario_bono = datetime.datetime.strptime(str(anio_actual-1)+'-07-01', '%Y-%m-%d').date()
                            dias_antiguo_salario_bono = (self.fecha_inicio - fecha_inicio_antiguo_salario_bono).days + 1

                            # cambiar salario por salario_anterior
                            bono_antiguo_salario = ((salario_anterior * 12)/ 365)*dias_antiguo_salario_bono
                            bono_anual = bono_nuevo_salario + bono_antiguo_salario
                            # ---------------------------- FIN CALCULO BONO

                            # ---------------------------- INICIO CALCULO AGUINALDO
                            fecha_final_nuevo_salario_aguinaldo = datetime.datetime.strptime(str(anio_actual)+'-11-30', '%Y-%m-%d').date()
                            dias_nuevo_salario_aguinaldo = (fecha_final_nuevo_salario_aguinaldo - self.fecha_inicio).days + 1
                            aguinaldo_nuevo_salario = ((salario * 12)/ 365)*dias_nuevo_salario_aguinaldo

                            fecha_inicio_antiguo_salario_aguinaldo = datetime.datetime.strptime(str(anio_actual-1)+'-12-01', '%Y-%m-%d').date()
                            dias_antiguo_salario_aguinaldo = (self.fecha_inicio - fecha_inicio_antiguo_salario_aguinaldo).days + 1
                            aguinaldo_antiguo_salario = ((salario_anterior * 12)/ 365)*dias_antiguo_salario_aguinaldo
                            aguinaldo_anual = aguinaldo_nuevo_salario + aguinaldo_antiguo_salario

                            # ---------------------------- FIN CALCULO AGUINALDO


                            fecha_final_calculo = datetime.datetime.strptime(str(anio_actual)+'-12-31', '%Y-%m-%d').date()
                            dias_trabajados = (fecha_final_calculo - self.fecha_inicio).days + 1
                            proyeccion_base_extra = ((empleado.contract_id.base_extra * 12) / 365)*dias_trabajados
                            otros_ingresos_gravados = empleado_planillas["otro_ingreso"] + proyeccion_base_extra

                            valor_salario_nuevo = ((salario * 12)/365)*dias_trabajados
                            # cuota_igss = (empleado_planillas["igss_total"] * -1) + (valor_salario_nuevo)
                            cuota_igss = (empleado_planillas["igss_total"] * -1) + (valor_salario_nuevo*0.0483)


                            # otra_info = self._get_informacion(empleado.id, '01-07-'+str(self.anio-1), '01-06-'+str(self.anio))
                            #hoja.write(fila, 2, fecha_actualizacion, formato_fecha)
                            hoja.write(fila, 2, renta_patrono_actual)
                            hoja.write(fila, 3, bono_anual)
                            hoja.write(fila, 4, aguinaldo_anual)
                            hoja.write(fila, 35, otros_ingresos_gravados)
                            hoja.write(fila, 36, aguinaldo_anual)
                            hoja.write(fila, 37, bono_anual)
                            hoja.write(fila, 38, cuota_igss)
                            hoja.write(fila, 39, fecha_actualizacion, formato_fecha)
                        fila += 1

        if self.tipo == "ajuste_suspesion":
            hoja_carga_ajuste = libro.add_worksheet('CargasAjustesysuspensiones')
            hoja_carga_ajuste.write(0, 0, 'NIT Empleado')
            hoja_carga_ajuste.write(0, 1, 'AJUSTE/SUSPENSION')

            retencion_pago = False
            if self.fecha_inicio and self.fecha_fin:
                retencion_pago = self._get_retencion_pago(self.env.context.get('active_ids', []), self.fecha_inicio, self.fecha_fin)
                fila = 1
                for empleado in self._get_empleados(self.env.context.get('active_ids', [])):
                    if empleado.id in retencion_pago and retencion_pago[empleado.id][2] < 0:
                        hoja_carga_ajuste.write(fila, 0, empleado.nit if empleado.nit else '')
                        hoja_carga_ajuste.write(fila, 1, retencion_pago[empleado.id][2] * -1)
                    fila += 1

        if self.tipo == "liquidacion_labor":
            hoja_fin_labores = libro.add_worksheet('CargaLiquidación Fin de labores')
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
                if empleado.contract_id.date_end and (empleado.contract_id.date_end >= self.fecha_inicio and empleado.contract_id.date_end <= self.fecha_fin):
                    hoja_fin_labores.write(fila, 0, empleado.nit if empleado.nit else '')
                    # hoja_fin_labores.write(fila, 3, (empleado.contract_id.wage))
                    hoja_fin_labores.write(fila, 4, empleado.contract_id.wage)
                    hoja_fin_labores.write(fila, 5, empleado.contract_id.wage)


                    otra_info = self._get_informacion(empleado.id, self.fecha_inicio, empleado.contract_id.date_end, 'liquidacion_labor')
                    hoja_fin_labores.write(fila, 1, otra_info['renta_patrono_actual'])
                    hoja_fin_labores.write(fila, 2, otra_info['bono_anual'])
                    hoja_fin_labores.write(fila, 3, otra_info['aguinaldo_anual'])
                    hoja_fin_labores.write(fila, 34, otra_info['otro_ingreso'])


                    hoja_fin_labores.write(fila, 38, otra_info['viaticos'])
                    hoja_fin_labores.write(fila, 39, otra_info['aguinaldo_anual'])
                    hoja_fin_labores.write(fila, 40, otra_info['bono_anual'])
                    hoja_fin_labores.write(fila, 41, otra_info['igss_total'])
                    hoja_fin_labores.write(fila, 42,  str(empleado.contract_id.date_end), formato_fecha)
                    fila += 1




        if self.tipo == "liquidacion_periodo":
            hoja_fin_periodo = libro.add_worksheet('CargaLiquidación Fin Período')
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
                # otra_info = self._get_informacion(empleado.id, '01-07-'+str(self.anio-1), '31-12-'+str(self.anio-1))
                if self.fecha_inicio and self.fecha_fin:
                    anio_actual = self.fecha_fin.year
                    otra_info = self._get_informacion(empleado.id, self.fecha_inicio, self.fecha_fin,'liquidacion_periodo')
                    hoja_fin_periodo.write(fila, 0, empleado.nit if empleado.nit else '')
                    hoja_fin_periodo.write(fila, 1, (empleado.contract_id.wage*2))
                    hoja_fin_periodo.write(fila, 2, otra_info['bono_anual'])
                    hoja_fin_periodo.write(fila, 3, otra_info['aguinaldo_anual'])

                    if empleado.contract_id.date_end and (empleado.contract_id.date_end > self.fecha_inicio and empleado.contract_id.date_end <= self.fecha_fin):
                        otra_info = self._get_informacion(empleado.id, '01-07-'+str(anio_actual), empleado.contract_id.date_end,'liquidacion_periodo')
                    else:
                        otra_info = self._get_informacion(empleado.id, self.fecha_inicio, self.fecha_fin,'liquidacion_periodo')

                    hoja_fin_periodo.write(fila, 34, otra_info['otro_ingreso'])
                    hoja_fin_periodo.write(fila, 38, otra_info['viaticos'])
                    hoja_fin_periodo.write(fila, 39, empleado.contract_id.wage)
                    hoja_fin_periodo.write(fila, 40, empleado.contract_id.wage)
                    hoja_fin_periodo.write(fila, 41, otra_info['igss_total'])
                fila += 1

        if self.tipo == "retencion_pago":
            hoja_retencion = libro.add_worksheet('Retención por pago')
            hoja_retencion.write(0, 0, 'NIT empleado')
            hoja_retencion.write(0, 1, 'Base Gravado')
            hoja_retencion.write(0, 2, 'Retencion por pago')
            hoja_retencion.write(0, 3, 'Fecha de retencion')
            fila = 1

            retencion_pago = self._get_retencion_pago(self.env.context.get('active_ids', []), self.fecha_inicio, self.fecha_fin)
            for empleado in self._get_empleados(self.env.context.get('active_ids', [])):
                if  retencion_pago and empleado.id in retencion_pago and retencion_pago[empleado.id][1] < 0:
                    hoja_retencion.write(fila, 0, empleado.nit if empleado.nit else '')
                    hoja_retencion.write(fila, 1, retencion_pago[empleado.id][0])
                    hoja_retencion.write(fila, 2, retencion_pago[empleado.id][1]*-1 if retencion_pago[empleado.id][1]<0 else retencion_pago[empleado.id][1] )
                    hoja_retencion.write(fila, 3, str(self.fecha_fin), formato_fecha)

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
