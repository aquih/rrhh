# -*- coding: utf-8 -*-

from odoo import models, fields, api
from odoo.release import version_info
import logging
import datetime
import time
import dateutil.parser
from dateutil.relativedelta import relativedelta
from dateutil import relativedelta as rdelta
from odoo.fields import Date, Datetime
from odoo.addons.l10n_gt_extra import a_letras

class HrPayslip(models.Model):
    _inherit = 'hr.payslip'

    porcentaje_prestamo = fields.Float(related="payslip_run_id.porcentaje_prestamo",string='Prestamo (%)',store=True)
    etiqueta_empleado_ids = fields.Many2many('hr.employee.category',string='Etiqueta empleado', related='employee_id.category_ids')
    cuenta_analitica_id = fields.Many2one('account.analytic.account','Cuenta analítica')

    # Dias trabajdas de los ultimos 12 meses hasta la fecha
    def dias_trabajados_ultimos_meses(self,empleado_id,fecha_desde,fecha_hasta):
        dias = {'days': 0}
        if empleado_id.contract_id.date_start:
            diferencia_meses = (fecha_hasta - fecha_desde)
            if empleado_id.contract_id.date_start <= fecha_hasta and empleado_id.contract_id.date_start >= fecha_desde:
                diferencia_meses = fecha_hasta - empleado_id.contract_id.date_start
        return diferencia_meses.days + 1

    def existe_entrada(self,entrada_ids,entrada_id):
        existe_entrada = False
        for entrada in entrada_ids:
            if entrada.input_type_id.id == entrada_id.id:
                existe_entrada = True
        return existe_entrada

    def compute_sheet(self):
        for nomina in self:
            mes_nomina = int(nomina.date_from.month)
            dia_nomina = int(nomina.date_to.day)
            anio_nomina = int(nomina.date_from.year)
            valor_pago = 0
            porcentaje_pagar = 0
            for entrada in nomina.input_line_ids:
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
                                porcentaje_pagar =(valor_pago * (nomina.porcentaje_prestamo/100))
                                entrada.amount = porcentaje_pagar
                        cantidad_pagos = prestamo.numero_descuentos
                        cantidad_pagados = 0
                        for lineas in prestamo.prestamo_ids:
                            if lineas.nomina_id:
                                cantidad_pagados +=1
                        if cantidad_pagados > 0 and cantidad_pagados < cantidad_pagos:
                            prestamo.estado = "proceso"
                        if cantidad_pagados == cantidad_pagos and cantidad_pagos > 0:
                            prestamo.estado = "pagado"
        res =  super(HrPayslip, self).compute_sheet()
        return res

    def _obtener_entrada(self,contrato_id):
        entradas = False
        if contrato_id.structure_type_id and contrato_id.structure_type_id.default_struct_id:
            if contrato_id.structure_type_id.default_struct_id.input_line_type_ids:
                entradas = [entrada for entrada in contrato_id.structure_type_id.default_struct_id.input_line_type_ids]
        return entradas

    def calculo_rrhh(self,nomina):
        salario = self.salario_promedio(self.employee_id,self.date_to)
        dias = self.dias_trabajados_ultimos_meses(self.contract_id.employee_id,self.date_from,self.date_to)
        for entrada in self.input_line_ids:
            if entrada.input_type_id.code == 'SalarioPromedio':
                entrada.amount = salario
            if entrada.input_type_id.code == 'DiasTrabajados12Meses':
                entrada.amount = dias
        return True

    # SALARIO PROMEDIO POR 12 MESES LABORADOS O MENOS
    def salario_promedio(self,empleado_id,fecha_final_nomina):
        historial_salario = []
        salario_meses = {}
        salario_total = 0
        extra_ordinario_total = 0
        salario_promedio_total = 0
        if empleado_id.contract_ids[0].historial_salario_ids:
            for linea in empleado_id.contract_ids[0].historial_salario_ids:
                historial_salario.append({'salario': linea.salario, 'fecha':linea.fecha})

            historial_salario_ordenado = sorted(historial_salario, key=lambda k: k['fecha'],reverse=True)
            meses_laborados = (fecha_final_nomina.year - empleado_id.contract_ids[0].date_start.year) * 12 + (fecha_final_nomina.month - empleado_id.contract_ids[0].date_start.month)

            contador_mes = 0
            if meses_laborados >= 12:
                while contador_mes < 12:
                    mes = relativedelta(months=contador_mes)
                    resta_mes = fecha_final_nomina - mes
                    mes_letras = a_letras.mes_a_letras(resta_mes.month-1)
                    llave = '01-'+str(resta_mes.month)+'-'+str(resta_mes.year)
                    salario_meses[llave] = {'nombre':mes_letras.upper(),'salario': 0,'anio':resta_mes.year,'extra':0,'total':0}
                    contador_mes += 1
            else:

                while contador_mes <= meses_laborados:
                    mes = relativedelta(months=contador_mes)
                    resta_mes = fecha_final_nomina - mes
                    mes_letras = a_letras.mes_a_letras(resta_mes.month-1)
                    llave = '01-'+str(resta_mes.month)+'-'+str(resta_mes.year)
                    salario_meses[llave] = {'nombre':mes_letras.upper(),'salario': 0,'anio':resta_mes.year,'extra':0,'total':0}
                    contador_mes += 1

            contador_mes = 0
            fecha_inicio_diferencia = datetime.datetime.strptime(str(historial_salario_ordenado[0]['fecha']), '%Y-%m-%d')
            # Sumamos + 1 en diferencia_meses por que no toma la fecha final
            diferencia_meses = ((fecha_final_nomina.year - fecha_inicio_diferencia.year) * 12 + (fecha_final_nomina.month - fecha_inicio_diferencia.month)) +1
            if len(historial_salario_ordenado) > 1:
                diferencia_meses = relativedelta(fecha_final_nomina, fecha_inicio_diferencia).months + 1
            for linea in historial_salario_ordenado:
                contador = 0
                while contador < diferencia_meses:

                    mes = relativedelta(months=contador_mes)
                    resta_mes = datetime.datetime.strptime(str(fecha_final_nomina),'%Y-%m-%d') - mes
                    mes_letras = a_letras.mes_a_letras(resta_mes.month-1)
                    llave = '01-'+str(resta_mes.month)+'-'+str(resta_mes.year)
                    if llave in salario_meses:
                        salario_meses[llave]['salario'] = linea['salario']
                        salario_total += linea['salario']
                    contador += 1
                    contador_mes += 1

                if len(historial_salario_ordenado) > 1:
                    fecha_cambio_salario = datetime.datetime.strptime(str(linea['fecha']), '%Y-%m-%d')

                    posicion_siguiente = historial_salario_ordenado.index(linea) + 1
                    if posicion_siguiente < len(historial_salario_ordenado):
                        fecha_inicio_diferencia = datetime.datetime.strptime(str(historial_salario_ordenado[posicion_siguiente]['fecha']), '%Y-%m-%d')
                        diferencia_meses = (fecha_cambio_salario.year - fecha_inicio_diferencia.year) * 12 + (fecha_cambio_salario.month - fecha_inicio_diferencia.month)

            salario_meses = sorted(salario_meses.items())
            salario_promedio_total =  (salario_total + extra_ordinario_total) / len(salario_meses)
        else:
            salario_promedio_total = empleado_id.contract_ids[0].wage
        return salario_promedio_total


    def horas_sumar(self,lineas):
        horas = 0
        dias = 0
        for linea in lineas:
            tipo_id = self.env['hr.work.entry.type'].search([('id','=',linea['work_entry_type_id'])])
            if tipo_id and tipo_id.is_leave and tipo_id.descontar_nomina == False:
                horas += linea['number_of_hours']
                dias += linea['number_of_days']
        return {'dias':dias, 'horas': horas}

    def _get_worked_day_lines(self):
        res = super(HrPayslip, self)._get_worked_day_lines()
        tipos_ausencias_ids = self.env['hr.leave.type'].search([])
        datos = self.horas_sumar(res)
        ausencias_restar = []

        dias_ausentados_restar = 0
        contracts = False
        if self.employee_id.contract_id:
            contracts = self.employee_id.contract_id

        for ausencia in tipos_ausencias_ids:
            if ausencia.work_entry_type_id and ausencia.work_entry_type_id.descontar_nomina:
                ausencias_restar.append(ausencia.work_entry_type_id.id)

        trabajo_id = self.env['hr.work.entry.type'].search([('code','=','TRABAJO100')])
        for r in res:
            tipo_id = self.env['hr.work.entry.type'].search([('id','=',r['work_entry_type_id'])])
            if tipo_id and tipo_id.is_leave == False:
                r['number_of_hours'] += datos['horas']
                r['number_of_days'] += datos['dias']

            if len(ausencias_restar)>0:
                if r['work_entry_type_id'] in ausencias_restar:
                    dias_ausentados_restar += r['number_of_days']

        if contracts:
            dias_laborados = 0
            if contracts.schedule_pay == 'monthly':
                dias_laborados = 30
            if contracts.schedule_pay == 'bi-monthly':
                dias_laborados = 15

            if contracts.date_start and self.date_from <= contracts.date_start <= self.date_to:
                dias_laborados = dias_laborados - (contracts.date_start - self.date_from).days
                res.append({'work_entry_type_id': trabajo_id.id, 'sequence': 10, 'number_of_days': dias_laborados - dias_ausentados_restar})
            else:
                if contracts.schedule_pay == 'monthly':
                    res.append({'work_entry_type_id': trabajo_id.id,'sequence': 10,'number_of_days': 30 - dias_ausentados_restar})
                if contracts.schedule_pay == 'bi-monthly':
                    res.append({'work_entry_type_id': trabajo_id.id,'sequence': 10,'number_of_days': 15 - dias_ausentados_restar})
                # Cálculo de días para catorcena
                if contracts.schedule_pay == 'bi-weekly':
                    dias_laborados = self.employee_id._get_work_days_data(Datetime.from_string(self.date_from), Datetime.from_string(self.date_to), calendar=contracts.resource_calendar_id)
                    res.append({'work_entry_type_id': trabajo_id.id,'sequence': 10,'number_of_days': (dias_laborados['days']+1 - dias_ausentados_restar)})
        return res

    @api.onchange('employee_id','struct_id','contract_id', 'date_from', 'date_to','porcentaje_prestamo')
    def _onchange_employee(self):
        res = super(HrPayslip, self)._onchange_employee()
        mes_nomina = self.date_from.month
        anio_nomina = self.date_from.year
        dia_nomina = self.date_to.day
        entradas_nomina = []
        if self.contract_id:
            entradas = self._obtener_entrada(self.contract_id)
            if self.contract_id.analytic_account_id:
                self.cuenta_analitica_id = self.contract_id.analytic_account_id.id
            if entradas:
                for entrada in entradas:
                    existe_entrada = False
                    if self.input_line_ids:
                        existe_entrada = self.existe_entrada(self.input_line_ids,entrada)
                    if existe_entrada == False:
                        entradas_nomina.append((0, 0, {'input_type_id':entrada.id}))
            if entradas_nomina:
                self.input_line_ids = entradas_nomina
            self.calculo_rrhh(self)

        for prestamo in self.employee_id.prestamo_ids:
            anio_prestamo = int(prestamo.fecha_inicio.year)
            for entrada in self.input_line_ids:
                if (prestamo.codigo == entrada.input_type_id.code) and ((prestamo.estado == 'nuevo') or (prestamo.estado == 'proceso')):
                    for lineas in prestamo.prestamo_ids:
                        if mes_nomina == int(lineas.mes) and anio_nomina == int(lineas.anio):
                            entrada.amount = lineas.monto*(self.porcentaje_prestamo/100)
        return res

class HrPayslipRun(models.Model):
    _inherit = 'hr.payslip.run'

    porcentaje_prestamo = fields.Float('Prestamo (%)')

    def generar_pagos(self):
        pagos = self.env['account.payment'].search([('nomina_id', '!=', False)])
        nominas_pagadas = []
        for pago in pagos:
            nominas_pagadas.append(pago.nomina_id.id)
        for nomina in self.slip_ids:
            if nomina.id not in nominas_pagadas:
                total_nomina = 0
                if nomina.employee_id.diario_pago_id and nomina.employee_id.address_home_id and nomina.state == 'done':
                    res = self.env['report.rrhh.recibo'].lineas(nomina)
                    total_nomina = res['totales'][0] + res['totales'][1]
                    pago = {
                        'payment_type': 'outbound',
                        'partner_type': 'supplier',
                        'payment_method_id': 2,
                        'partner_id': nomina.employee_id.address_home_id.id,
                        'amount': total_nomina,
                        'journal_id': nomina.employee_id.diario_pago_id.id,
                        'nomina_id': nomina.id
                    }
                    pago_id = self.env['account.payment'].create(pago)
        return True
