# -*- coding: utf-8 -*-

from odoo import models, fields, api, _
import logging
import datetime
import time
import dateutil.parser
from dateutil.relativedelta import relativedelta
from dateutil import relativedelta as rdelta
from odoo.fields import Date, Datetime
from calendar import monthrange
from odoo.addons.l10n_gt_extra import a_letras
from odoo.exceptions import ValidationError

class HrPayslip(models.Model):
    _inherit = 'hr.payslip'

    porcentaje_prestamo = fields.Float(related="payslip_run_id.porcentaje_prestamo", string='Prestamo (%)', store=True)
    cuenta_analitica_id = fields.Many2one('account.analytic.account', 'Cuenta analítica')
    descuento_isr = fields.Boolean(related="payslip_run_id.descuento_isr", string='Descuento ISR', store=True)
    
    # TODO: Quitar en la siguiente versión
    etiqueta_empleado_ids = fields.Many2many('hr.employee.category', string='Etiqueta empleado', related='employee_id.category_ids') # no parece usarse

    def dias_trabajados_rango(self, empleado_id, fecha_desde, fecha_hasta):
        if empleado_id.date_start:
            diferencia_meses = (fecha_hasta - fecha_desde)
            if fecha_desde <= empleado_id.date_start <= fecha_hasta:
                diferencia_meses = fecha_hasta - empleado_id.date_start

        return diferencia_meses.days + 1

    def calcular_sueldo_devengado(self, nomina):
        fecha_inicio = nomina.date_to.replace(month=1, day=1)
        fecha_fin = nomina.date_to
        nomina_ids = self.env['hr.payslip'].search([('employee_id','=', nomina.employee_id.id), ('date_from', '>=', fecha_inicio), ('date_to', '<=', fecha_fin)])

        devengado = 0
        for n in nomina_ids:
            for linea in n.line_ids:
                if linea.salary_rule_id.id in n.employee_id.company_id.salario_ids.ids:
                    devengado += linea.total
        return devengado

    def calcular_sueldo_proyectado(self, nomina):
        fecha_inicio = nomina.date_to.replace(month=1, day=1)
        fecha_fin_proyectar = nomina.date_to.replace(month=12, day=31)
        meses_proyectar = (fecha_fin_proyectar.month - nomina.date_to.month)
        nomina_ids = self.env['hr.payslip'].search([('employee_id','=', nomina.employee_id.id), ('date_from', '>=', fecha_inicio), ('date_to', '<=',  nomina.date_to)], limit=1, order='date_to desc')

        proyectado = 0
        ultimo_salario = 0
        for n in nomina_ids:
            for linea in n.line_ids:
                if linea.salary_rule_id.id in n.employee_id.company_id.salario_total_ids.ids:
                    ultimo_salario += linea.total

        return ultimo_salario * meses_proyectar

    def calcular_sueldos(self, nomina):
        devengado = self.calcular_sueldo_devengado(nomina)
        proyectado = self.calcular_sueldo_proyectado(nomina)

        return devengado + proyectado

    def calcular_horas_extras(self, nomina):
        fecha_inicio = nomina.date_to.replace(month=1, day=1)
        fecha_fin = nomina.date_to
        fecha_fin_proyectar = nomina.date_to.replace(month=12, day=31)
        meses_proyectar = (fecha_fin_proyectar.month - nomina.date_to.month)
        meses_transcurrido = (nomina.date_to.month - fecha_inicio.month) + 1
        nomina_ids = self.env['hr.payslip'].search([('employee_id','=', nomina.employee_id.id), ('date_from', '>=', fecha_inicio), ('date_to', '<=', fecha_fin)])

        horas_extras = 0
        for n in nomina_ids:
            for linea in n.line_ids:
                if linea.salary_rule_id.id in n.employee_id.company_id.horas_extras_ids.ids:
                    horas_extras += linea.total
        
        return horas_extras_devengado + ((horas_extras / meses_transcurrido ) * meses_proyectar)

    def calcular_bonificacion_decreto_devengado(self, nomina):
        fecha_inicio = nomina.date_to.replace(month=1, day=1)
        fecha_fin = nomina.date_to
        nomina_ids = self.env['hr.payslip'].search([('employee_id','=', nomina.employee_id.id), ('date_from', '>=', fecha_inicio), ('date_to', '<=', fecha_fin)])

        devengado = 0
        for n in nomina_ids:
            for linea in n.line_ids:
                if linea.salary_rule_id.id in n.employee_id.company_id.bonificacion_ids.ids:
                    devengado += linea.total

        return devengado

    def calcular_bonificacion_decreto_proyectado(self, nomina):
        fecha_inicio = nomina.date_to.replace(month=1, day=1)
        fecha_fin_proyectar = nomina.date_to.replace(month=12, day=31)
        meses_proyectar = (fecha_fin_proyectar.month - nomina.date_to.month)
        nomina_ids = self.env['hr.payslip'].search([('employee_id','=', nomina.employee_id.id), ('date_from', '>=', fecha_inicio), ('date_to', '<=',  nomina.date_to)], limit=1, order='date_to desc')

        ultimo_salario = 0
        for n in nomina_ids:
            for linea in n.line_ids:
                if linea.salary_rule_id.id in n.employee_id.company_id.bonificacion_ids.ids:
                    ultimo_salario += linea.total

        return ultimo_salario * meses_proyectar

    def calcular_bonificacion_decreto(self, nomina):
        bonificacion_decreto = 0
        devengado = self.calcular_bonificacion_decreto_devengado(nomina)
        proyectado = self.calcular_bonificacion_decreto_proyectado(nomina)

        return devengado + proyectado

    def calcular_aguinaldo(self, nomina):
        aguinaldo = (nomina.wage * 12) / 12
        if nomina.company_id.isr_sueldo_base_extra:
            aguinaldo = ((nomina.wage + nomina.base_extra) * 12) / 12

        return aguinaldo

    def calcular_bonoc(self, nomina):
        bonoc = (nomina.wage * 12) / 12
        if nomina.company_id.isr_sueldo_base_extra:
            bonoc = ((nomina.wage + nomina.base_extra) * 12) / 12
            
        return bonoc

    def calcular_otro_ingreso_afecto(self, nomina):
        fecha_inicio = nomina.date_to.replace(month=1, day=1)
        fecha_fin = nomina.date_to
        nomina_ids = self.env['hr.payslip'].search([('employee_id','=', nomina.employee_id.id), ('date_from', '>=', fecha_inicio), ('date_to', '<=', fecha_fin)])

        otro_ingreso = 0
        for n in nomina_ids:
            for linea in n.line_ids:
                if linea.salary_rule_id.id in n.employee_id.company_id.otro_ingreso_afecto_ids.ids:
                    otro_ingreso += linea.total
                    
        return otro_ingreso

    def calcular_bono_productividad(self, nomina):
        fecha_inicio = nomina.date_to.replace(month=1, day=1)
        fecha_fin = nomina.date_to
        nomina_ids = self.env['hr.payslip'].search([('employee_id','=', nomina.employee_id.id),('date_from', '>=', fecha_inicio),('date_to', '<=', fecha_fin)])

        bono_productividad = 0
        for n in nomina_ids:
            for linea in n.line_ids:
                if linea.salary_rule_id.id in n.employee_id.company_id.bonificaciones_adicionales_ids.ids:
                    bono_productividad += linea.total

        if nomina.company_id.proyectar_bonificaciones_adicionales:
            bono_productividad += self.calcular_bono_productividad_proyectado(nomina)

        return bono_productividad

    def calcular_bono_productividad_proyectado(self, nomina):
        fecha_inicio = nomina.date_to.replace(month=1, day=1)
        fecha_fin_proyectar = nomina.date_to.replace(month=12, day=31)
        meses_proyectar = (fecha_fin_proyectar.month - nomina.date_to.month)
        nomina_ids = self.env['hr.payslip'].search([('employee_id','=', nomina.employee_id.id), ('date_from', '>=', fecha_inicio), ('date_to', '<=',  nomina.date_to)])

        bono_productividad = 0
        for n in nomina_ids:
            for linea in n.line_ids:
                if linea.salary_rule_id.id in n.employee_id.company_id.bonificaciones_adicionales_ids.ids:
                    bono_productividad += linea.total

        return bono_productividad * meses_proyectar

    def calcular_igss_devengado(self, nomina):
        fecha_inicio = nomina.date_to.replace(month=1, day=1)
        nomina_ids = self.env['hr.payslip'].search([('employee_id','=', nomina.employee_id.id), ('date_from', '>=', fecha_inicio), ('date_to', '<=', fecha_fin)])

        igss_devengado = 0
        for n in nomina_ids:
            for linea in n.line_ids:
                if linea.salary_rule_id.id in n.employee_id.company_id.igss_ids.ids:
                    igss_devengado += linea.total

        return igss_devengado

    def calcular_igss_proyectado(self, nomina):
        fecha_inicio = nomina.date_to.replace(month=1, day=1)
        fecha_fin_proyectar = nomina.date_to.replace(month=12, day=31)
        meses_proyectar = (fecha_fin_proyectar.month - nomina.date_to.month)
        nomina_ids = self.env['hr.payslip'].search([('employee_id','=', nomina.employee_id.id), ('date_from', '>=', fecha_inicio), ('date_to', '<=',  nomina.date_to)])

        igss_proyectado = 0
        for n in nomina_ids:
            for linea in n.line_ids:
                if linea.salary_rule_id.id in n.employee_id.company_id.igss_ids.ids:
                    igss_proyectado += linea.total

        return igss_proyectado * meses_proyectar

    def calcular_cuota_igss(self, nomina):
        igss_devengado = self.calcular_igss_devengado(nomina)
        igss_proyectado = self.calcular_igss_proyectado(nomina)

        return igss_devengado + igss_proyectado

    def calcular_retencion_isr_descontado(self, nomina):
        fecha_inicio = nomina.date_to.replace(month=1, day=1)
        fecha_fin = nomina.date_to
        nomina_ids = self.env['hr.payslip'].search([('employee_id','=', nomina.employee_id.id), ('date_from', '>=', fecha_inicio), ('date_to', '<', fecha_fin)])

        isr_descontado = 0
        for n in nomina_ids:
            for linea in n.line_ids:
                if linea.salary_rule_id.id in n.employee_id.company_id.isr_ids.ids:
                    isr_descontado += linea.total

        return isr_descontado

    def ajuste_isr(self, nomina):
        nomina_ids = self.env['hr.payslip'].search([('employee_id','=', nomina.employee_id.id), ('date_from', '>=', nomina.date_from), ('date_to', '<=', nomina.date_to)])

        ajuste = 0
        for n in nomina_ids:
            for linea in n.line_ids:
                if linea.salary_rule_id.id in n.employee_id.company_id.ajuste_ids.ids:
                    ajuste += (linea.total * -1) if linea.total < 0 else ( linea.total * -1 if linea.total > 0 else 0)

        return ajuste

    def calculo_isr(self, nomina):
        fecha_fin_proyectar = nomina.date_to.replace(month=12, day=31)
        meses_proyectar = (fecha_fin_proyectar.month - nomina.date_to.month)

        sueldos = self.calcular_sueldos(nomina)
        horas_extras = self.calcular_horas_extras(nomina)
        bonificacion_decreto = self.calcular_bonificacion_decreto(nomina)
        aguinaldo = self.calcular_aguinaldo(nomina)
        bonoc = self.calcular_bonoc(nomina)
        bono_productividad = self.calcular_bono_productividad(nomina)
        otro_ingreso_afecto = self.calcular_otro_ingreso_afecto(nomina)
        cuota_igss = abs(self.calcular_cuota_igss(nomina))
        ajuste = self.ajuste_isr(nomina)

        rubro_ingresos = sueldos + horas_extras + bonificacion_decreto + aguinaldo + bonoc + bono_productividad + otro_ingreso_afecto
        rubro_deducciones = aguinaldo + bonoc + abs(cuota_igss)
        deduccion_fija = nomina.company_id.monto_deduccion_fija
        renta_impunible = 0 if (rubro_ingresos - rubro_deducciones - deduccion_fija) < 0 else (rubro_ingresos - rubro_deducciones - deduccion_fija)
        rubro_renta_cinco = 15000 if ((renta_impunible * 0.05) > 15000) else (renta_impunible * 0.05)
        rubro_renta_siete = ((renta_impunible - 300000) * 0.07) if ((renta_impunible * 0.05) > 15000) else 0
        rubro_retencion_anual = rubro_renta_cinco + rubro_renta_siete
        retencion_isr_descontado = self.calcular_retencion_isr_descontado(nomina)
        retencion_isr_descontado_total = ((rubro_retencion_anual + retencion_isr_descontado) /  (meses_proyectar + 1) ) + ajuste
        isr_total = retencion_isr_descontado_total

        return {
            "sueldos": sueldos,
            "horas_extras": horas_extras,
            "bonificacion_decreto": bonificacion_decreto,
            "aguinaldo": aguinaldo,
            "bonoc": bonoc,
            "bono_productividad": bono_productividad,
            "otro_ingreso_afecto": otro_ingreso_afecto,
            "aguinaldo_mes": aguinaldo,
            "bonoc_mes": bonoc,
            "cuota_igss": cuota_igss,
            "rubro_ingresos": rubro_ingresos,
            "rubro_deducciones": rubro_deducciones,
            "deduccion_fija": deduccion_fija,
            "renta_impunible": renta_impunible,
            "rubro_renta_cinco": rubro_renta_cinco,
            "rubro_renta_siete": rubro_renta_siete,
            "rubro_retencion_anual": rubro_retencion_anual,
            "retencion_isr_descontado": retencion_isr_descontado_total,
            "isr_total": max(isr_total, 0),
        }

    def calculo_entradas_anuales(self, nomina):
        salario = self.salario_promedio(self.employee_id, self.date_to)
        dias = self.dias_trabajados_rango(self.employee_id, self.date_from,self.date_to)
        dias_calendario = monthrange(self.date_to.year, self.date_to.month)[1]

        for entrada in self.input_line_ids:
            if entrada.input_type_id.code == 'SalarioPromedio':
                entrada.amount = salario
            if entrada.input_type_id.code == 'DiasTrabajados12Meses':
                entrada.amount = dias
            if entrada.input_type_id.code == 'DiasCalendario':
                entrada.amount = dias_calendario
                
        return True

    # TODO: Borrar lo antes posible, ya no sirve para nada
    def salario_promedio_viejo(self, empleado_id, fecha_final_nomina):
        historial_salario = []
        salario_meses = {}
        salario_total = 0
        salario_promedio_total = 0
        extra_ordinario_total = 0
        salario_completo = {}
        salario_sumatoria = 0

        if empleado_id.historial_salario_ids:
            posicion_historial = 0
            for linea in empleado_id.historial_salario_ids:
                if (posicion_historial + 1) <= len(empleado_id.historial_salario_ids):
                    historial_salario.append({'salario': linea.salario, 'fecha':linea.fecha})
                    contador_mes_historial = 0

                    # Ordernar fechas cuando la fecha del historial no empieza en el primer día del mes
                    llave_salario = '01-'+str(linea.fecha.month)+'-'+str(linea.fecha.year)
                    llave_salario_fecha = datetime.datetime.strptime(str(llave_salario),'%d-%m-%Y').date()
                    llave_salario_fecha_str = '01-'+str(llave_salario_fecha.month)+'-'+str(llave_salario_fecha.year)
                    if posicion_historial+1 >= len(empleado_id.historial_salario_ids):
                        while llave_salario_fecha < fecha_final_nomina:
                            # Este código es el mismo de abajo, sería mejor refactorizar
                            salario_completo[str(llave_salario_fecha_str)] = linea.salario
                            mes = relativedelta(months=1)
                            llave_salario_fecha = llave_salario_fecha + mes
                            llave_salario_fecha_str = '01-'+str(llave_salario_fecha.month)+'-'+str(llave_salario_fecha.year)
                            contador_mes_historial += 1

                        posicion_historial += 1
                    else:
                        while llave_salario_fecha < empleado_id.historial_salario_ids[posicion_historial+1].fecha:
                            salario_completo[str(llave_salario_fecha_str)] = linea.salario
                            mes = relativedelta(months=1)
                            llave_salario_fecha = llave_salario_fecha + mes
                            llave_salario_fecha_str = '01-'+str(llave_salario_fecha.month)+'-'+str(llave_salario_fecha.year)
                            contador_mes_historial += 1

                        posicion_historial += 1
        
            fecha_inicio_contrato = datetime.datetime.strptime(str(empleado_id.date_start),"%Y-%m-%d")
            fecha_final_contrato = datetime.datetime.strptime(str(fecha_final_nomina),"%Y-%m-%d")
            meses_laborados = (fecha_final_contrato.year - fecha_inicio_contrato.year) * 12 + (fecha_final_contrato.month - fecha_inicio_contrato.month)

            contador_mes = 0
            if meses_laborados >= 12:
                while contador_mes < 12:
                    # Este código es el mismo de abajo, sería mejor refactorizar
                    mes = relativedelta(months=contador_mes)
                    resta_mes = fecha_final_contrato - mes
                    mes_letras = a_letras.mes_a_letras(resta_mes.month-1)
                    llave = '01-'+str(resta_mes.month)+'-'+str(resta_mes.year)
                    salario = 0
                    if llave in salario_completo:
                        salario = salario_completo[llave]
                    salario_meses[llave] = {'nombre':mes_letras.upper(),'salario': salario,'anio':resta_mes.year,'extra':0,'total':0}
                    contador_mes += 1
                    salario_sumatoria += salario
            else:
                while contador_mes <= meses_laborados:
                    mes = relativedelta(months=contador_mes)
                    resta_mes = fecha_final_contrato - mes
                    mes_letras = a_letras.mes_a_letras(resta_mes.month-1)
                    llave = '01-'+str(resta_mes.month)+'-'+str(resta_mes.year)
                    salario = 0
                    if llave in salario_completo:
                        salario = salario_completo[llave]
                    salario_meses[llave] = {'nombre':mes_letras.upper(),'salario': salario,'anio':resta_mes.year,'extra':0,'total':0}
                    contador_mes += 1
                    salario_sumatoria += salario

            salario_promedio_total =  salario_sumatoria / len(salario_meses)
        else:
            salario_promedio_total = empleado_id.wage

        return salario_promedio_total

    # Salario promedio por 12 meses laborados o menos si el contrato empezó antes
    def salario_promedio(self, empleado_id, fecha_final_nomina):
        salario_por_mes = []
        if empleado_id.historial_salario_ids:
            for mes_restar in range(12):
                fecha = fecha_final_nomina - relativedelta(months=mes_restar)

                # Ya no se continua el ciclo si el empleado a laborado menos de 12 meses
                if empleado_id.date_start <= fecha:
                    anio = fecha.strftime('%Y')
                    mes = fecha.strftime('%m')

                    # Se busca el salario mas reciente que sea menor o igual al del mes que se está recorriendo
                    salario = 0
                    historial = empleado_id.historial_salario_ids.filtered(lambda h: f'{h.anio}-{str(h.mes).zfill(2)}' <= f'{anio}-{mes}').sorted(key=lambda h: f'{h.anio}-{h.mes}', reverse=True)
                    if historial:
                        salario = historial[0].salario
                    
                    salario_por_mes.append({'mes': f'{anio}-{mes}', 'salario': salario})
    
        salario_promedio_total = 0
        if len(salario_por_mes) > 0:
            salario_promedio_total = sum([s['salario'] for s in salario_por_mes])/len(salario_por_mes)
        
        return salario_promedio_total

    def dias_horas_sumar(self, lineas):
        horas = 0
        dias = 0
        for linea in lineas:
            tipo_id = self.env['hr.work.entry.type'].browse(linea['work_entry_type_id'])
            if tipo_id and tipo_id.is_leave and tipo_id.descontar_nomina == False:
                horas += linea['number_of_hours']
                dias += linea['number_of_days']

        return {'dias': dias, 'horas': horas}

    # TODO: dejar de usar, odoo 19.0 ya debería de poder funcionar sin necesidad de estos cálculos
    def _get_worked_day_lines(self, domain=None, check_out_of_version=True):
        res = super()._get_worked_day_lines(domain=None, check_out_of_version=True)
        tipos_ausencias_ids = self.env['hr.leave.type'].search([])
        datos = self.dias_horas_sumar(res)
        ausencias_restar = []

        dias_ausentados_restar = 0

        for ausencia in tipos_ausencias_ids:
            if ausencia.work_entry_type_id and ausencia.work_entry_type_id.descontar_nomina:
                ausencias_restar.append(ausencia.work_entry_type_id.id)

        trabajo_id = self.env['hr.work.entry.type'].search([('code','=','TRABAJO100')])
        for r in res:
            tipo_id = self.env['hr.work.entry.type'].browse(r['work_entry_type_id'])
            if tipo_id and not tipo_id.is_leave:
                r['number_of_hours'] += datos['horas']
                r['number_of_days'] += datos['dias']

            if len(ausencias_restar) > 0:
                if r['work_entry_type_id'] in ausencias_restar:
                    dias_ausentados_restar += r['number_of_days']

        dias_laborados = 0
        if self.struct_id:
            if self.struct_id.schedule_pay == 'monthly':
                dias_laborados = 30
            if self.struct_id.schedule_pay == 'semi-monthly':
                dias_laborados = 15

        reference_calendar = self.employee_id.resource_calendar_id

        # Para determinar si la planilla es normal o anual (de aguinaldo o bono 14)
        dias_bonificacion = reference_calendar.get_work_duration_data(datetime.datetime.combine(self.date_from, datetime.datetime.min.time()), datetime.datetime.combine(self.date_to, datetime.datetime.max.time()), compute_leaves=False, domain=[])

        # Cuando es una planilla normal y de un empleado que ingresó después de la fecha de inicio la planilla
        if self.employee_id.date_start and dias_bonificacion['days'] <= 31 and self.date_from <= self.employee_id.date_start <= self.date_to:
            dias_laborados = dias_laborados - ((self.employee_id.date_start - self.date_from).days)

            # Cuando es una planilla normal, y el empleado entra y sale el mismo mes
            if self.employee_id.date_end and (self.date_from <= self.employee_id.date_end <= self.date_to):
                dias_laborados = ((self.employee_id.date_end - self.employee_id.date_start).days) + 1

            res.append({'work_entry_type_id': trabajo_id.id, 'sequence': 10, 'number_of_days': dias_laborados - dias_ausentados_restar})

        # Cuando es una planilla normal y de un empleado que salió antes de la fecha de fin de la planilla
        elif self.employee_id.date_end and dias_bonificacion['days'] <= 31 and self.date_from <= self.employee_id.date_end <= self.date_to:
            dias_laborados =  ((self.employee_id.date_end - self.date_from).days) + 1

            # Cuando el contrato finaliza dentro del rango en el que se genera la planilla, es necesario verificar si el pago es quincenal o mensual
            # por que necesitamos parametrizar que los dias trabajados no sea mayor que a los días dentro del rango de la planilla
            if self.employee_id.schedule_pay == 'semi-monthly':
                res.append({'work_entry_type_id': trabajo_id.id, 'sequence': 10, 'number_of_days': min(dias_laborados, 15) - dias_ausentados_restar})
            else:
                res.append({'work_entry_type_id': trabajo_id.id, 'sequence': 10, 'number_of_days': min(dias_laborados, 30) - dias_ausentados_restar})

        # Cuando es una planilla anual y de un empleado que ingresó antes de la fecha de inicio de la planilla
        elif dias_bonificacion['days'] > 150 and self.date_from >= self.employee_id.date_start:
            res.append({'work_entry_type_id': trabajo_id.id, 'sequence': 10, 'number_of_days': dias_bonificacion['days']+1})

        # Cuando es una planilla anual y de un empleado que ingresó después de la fecha de inicio de la planilla
        elif dias_bonificacion['days'] > 150 and self.date_from <= self.employee_id.date_start <= self.date_to:
            res.append({'work_entry_type_id': trabajo_id.id, 'sequence': 10, 'number_of_days': dias_bonificacion['days']+1})

        # Cuando el empleado ingreso antes de la fecha de la planilla y no ha salido
        else:
            # Cálculo para mensualidad
            if self.struct_id.schedule_pay == 'monthly' or self.employee_id.structure_type_id.default_schedule_pay == 'monthly':
                total_dias = 30 - dias_ausentados_restar
                res.append({'work_entry_type_id': trabajo_id.id,'sequence': 10,'number_of_days': 0 if total_dias < 0 else total_dias})

            # Cálculo para quincena
            if self.struct_id.schedule_pay == 'semi-monthly' or self.employee_id.structure_type_id.default_schedule_pay == 'semi-monthly':
                dias_periodo = min((self.date_to - self.date_from).days + 1, 15)

                # Calcular días de ausencia solo dentro del período de la nómina
                dias_ausencia_en_periodo = 0
                ausencias = self.env['hr.leave'].search([
                    ('employee_id', '=', self.employee_id.id),
                    ('state', '=', 'validate'),
                    ('request_date_from', '<=', self.date_to),
                    ('request_date_to', '>=', self.date_from),
                ])

                for ausencia in ausencias:
                    if ausencia.holiday_status_id.work_entry_type_id.descontar_nomina == True:
                        # Determinar el rango de intersección de la ausencia
                        inicio = max(ausencia.request_date_from, self.date_from)
                        fin = min(ausencia.request_date_to, self.date_to)
                        if inicio <= fin:
                            dias_ausencia_en_periodo += (fin - inicio).days + 1

                # Ahora calcular los días trabajados
                if dias_ausencia_en_periodo == 0:
                    dias_trabajados = 15
                elif dias_ausencia_en_periodo >= dias_periodo:
                    dias_trabajados = 0
                else:
                    dias_trabajados = dias_periodo - dias_ausencia_en_periodo

                res.append({ 'work_entry_type_id': trabajo_id.id, 'sequence': 10, 'number_of_days': dias_trabajados })

            # Cálculo de días para catorcena
            if self.struct_id.schedule_pay == 'weekly' or self.employee_id.structure_type_id.default_schedule_pay == 'weekly':
                dias_laborados = reference_calendar.get_work_duration_data(self.date_from, self.date_to, compute_leaves=False,domain = [])
                res.append({'work_entry_type_id': trabajo_id.id, 'sequence': 10, 'number_of_days': (dias_laborados['days'] + 1 - dias_ausentados_restar)})

        self.calculo_entradas_anuales(self)
        
        return res

    # TODO: dejar de usar: las entradas se pueden poner en el empleado y los
    # prestamos se pueden hacer con ajustes de salario
    @api.depends('employee_id', 'version_id', 'struct_id', 'date_from', 'date_to', 'struct_id')
    def _compute_input_line_ids(self):
        res = super()._compute_input_line_ids()
        for slip in self:
            if slip.employee_id and slip.struct_id and slip.struct_id.input_line_type_ids:
                input_line_vals = []
                if slip.input_line_ids:
                    slip.input_line_ids.unlink()

                for line in slip.struct_id.input_line_type_ids:
                    input_line_vals.append((0,0,{
                        'name': line.name,
                        'amount': 0,
                        'input_type_id': line.id,
                    }))
                slip.update({'input_line_ids': input_line_vals})

                mes_nomina = slip.date_from.month
                anio_nomina = slip.date_from.year
                dia_nomina = slip.date_to.day
                entradas_nomina = []
                if slip.employee_id.prestamo_ids:
                    for prestamo in slip.employee_id.prestamo_ids:
                        anio_prestamo = int(prestamo.fecha_inicio.year)
                        for entrada in slip.input_line_ids:
                            if (prestamo.codigo == entrada.input_type_id.code) and ((prestamo.estado == 'nuevo') or (prestamo.estado == 'proceso')):
                                valor_entrada = entrada.amount
                                for lineas in prestamo.prestamo_ids:
                                    if mes_nomina == int(lineas.mes) and anio_nomina == int(lineas.anio):
                                        valor_entrada += lineas.monto*(slip.porcentaje_prestamo/100)
                                entrada.amount = valor_entrada
        return res

    # TODO: dejar de usar: las entradas se pueden poner en el empleado y los
    # prestamos se pueden hacer con ajustes de salario
    def compute_sheet(self):
        for nomina in self:
            mes_nomina = int(nomina.date_from.month)
            dia_nomina = int(nomina.date_to.day)
            anio_nomina = int(nomina.date_from.year)
            for entrada in nomina.input_line_ids:
                valor_entrada = 0
                for prestamo in nomina.employee_id.prestamo_ids:
                    anio_prestamo = int(prestamo.fecha_inicio.year)
                    if (prestamo.codigo == entrada.input_type_id.code) and ((prestamo.estado == 'nuevo') or (prestamo.estado == 'proceso')):
                        lista = []
                        for lineas in prestamo.prestamo_ids:
                            if mes_nomina == int(lineas.mes) and anio_nomina == int(lineas.anio):
                                lista = lineas.nomina_id.ids
                                lista.append(nomina.id)
                                lineas.nomina_id = [(6, 0, lista)]
                                valor_pago = lineas.monto
                                valor_entrada +=(valor_pago * (nomina.porcentaje_prestamo/100))
                                entrada.amount = valor_entrada
                        cantidad_pagos = prestamo.numero_descuentos
                        cantidad_pagados = 0
                        for lineas in prestamo.prestamo_ids:
                            if lineas.nomina_id:
                                cantidad_pagados +=1
                        if cantidad_pagados > 0 and cantidad_pagados < cantidad_pagos:
                            prestamo.estado = "proceso"
                        if prestamo.pendiente_pagar_prestamo == 0:
                            prestamo.estado = "pagado"
        res =  super().compute_sheet()
        for nomina in self:
            if nomina.descuento_isr:
                calculos_isr = self.calculo_isr(nomina)
                for entrada in self.input_line_ids:
                    if entrada.input_type_id.code in calculos_isr:
                        entrada.amount = calculos_isr[entrada.input_type_id.code]
                nomina.with_context(payslip_no_recompute=True)._compute_line_ids()
        return res

    def action_payslip_cancel(self):
        for nomina in self:
            pagos_id = self.env['account.payment'].search([('nomina_id','=',nomina.id), ('state','=','posted')])
            if len(pagos_id) > 0:
                raise ValidationError(_("No puede cancelar por que tiene un pago asociado."))

        return super().action_payslip_cancel()

class HrPayslipRun(models.Model):
    _inherit = 'hr.payslip.run'

    porcentaje_prestamo = fields.Float('Prestamo (%)')
    descuento_isr = fields.Boolean('Descuento ISR')

    def generar_pagos(self):
        pagos = self.env['account.payment'].search([('nomina_id', '!=', False)])
        nominas_pagadas = []

        for pago in pagos:
            nominas_pagadas.append(pago.nomina_id.id)

        for nomina in self.slip_ids:
            if nomina.id not in nominas_pagadas:
                total_nomina = 0
                if nomina.employee_id.diario_pago_id and nomina.employee_id._get_related_partners() and nomina.state in ['done','paid']:
                    res = self.env['report.rrhh.recibo'].lineas(nomina)
                    total_nomina = res['totales'][0] + res['totales'][1]
                    payment_method_line_id = self.env["account.payment.method.line"].search([('journal_id','=', nomina.employee_id.diario_pago_id.id)])
                    
                    pago = {
                        'payment_type': 'outbound',
                        'partner_type': 'supplier',
                        'payment_method_line_id': payment_method_line_id[0].id,
                        'partner_id': nomina.employee_id._get_related_partners()[0].id,
                        'amount': total_nomina,
                        'journal_id': nomina.employee_id.diario_pago_id.id,
                        'nomina_id': nomina.id
                    }
                    pago_id = self.env['account.payment'].create(pago)

        return True
