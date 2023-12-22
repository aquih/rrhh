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
from calendar import monthrange
import odoo.addons.l10n_gt_extra.a_letras as a_letras

class HrPayslip(models.Model):
    _inherit = 'hr.payslip'

    porcentaje_prestamo = fields.Float(related="payslip_run_id.porcentaje_prestamo",string='Prestamo (%)',store=True)
    etiqueta_empleado_ids = fields.Many2many('hr.employee.category',string='Etiqueta empleado', related='employee_id.category_ids')

    @api.multi
    def onchange_employee_id(self, date_from, date_to, employee_id=False, contract_id=False):
        res = super(HrPayslip, self).onchange_employee_id(date_from,date_to,employee_id, contract_id=False)
        payslip_run_id = self.env['hr.payslip.run'].search([('id','=',self.env.context.get('active_id'))])
        if payslip_run_id and payslip_run_id.estructura_id:
            res['value']['struct_id'] = payslip_run_id.estructura_id.id
        return res

    @api.multi
    def action_payslip_done(self):
        res = super(HrPayslip, self).action_payslip_done()
        for slip in self:
            if slip.move_id:
                slip.move_id.button_cancel()
                for line in slip.move_id.line_ids:
                    line.analytic_account_id = slip.contract_id.analytic_account_id.id
                slip.move_id.post()
        return res

    # Dias trabajdas de los ultimos 12 meses hasta la fecha
    def dias_trabajados_ultimos_meses(self,empleado_id,fecha_desde,fecha_hasta):
        dias = {'days': 0}
        if empleado_id.contract_id.date_start:
            fecha_nomina_desde = datetime.datetime.strptime(str(fecha_hasta), '%Y-%m-%d').date()
            fecha_nomina_hasta = datetime.datetime.strptime(str(fecha_desde), '%Y-%m-%d').date()
            # diferencia_meses = relativedelta(fecha_nomina_desde,fecha_nomina_hasta)
            diferencia_meses = (fecha_nomina_desde - fecha_nomina_hasta)
            if empleado_id.contract_id.date_start <= fecha_hasta and empleado_id.contract_id.date_start >= fecha_desde:
                fecha_contrato_inicio = datetime.datetime.strptime(str(empleado_id.contract_id.date_start), '%Y-%m-%d').date()
                fecha_nomina_hasta = datetime.datetime.strptime(str(fecha_hasta), '%Y-%m-%d').date()
                diferencia_meses = fecha_nomina_hasta - fecha_contrato_inicio
            # if int(diferencia_meses.years) == 0:
            #     dias = empleado_id.get_work_days_data(Datetime.from_string(fecha_desde), Datetime.from_string(fecha_hasta), calendar=empleado_id.contract_id.resource_calendar_id)
            # else:
            #     mes = relativedelta(months=12)
            #     fecha_inicio = datetime.datetime.strptime(str(fecha_nomina_desde - mes), '%Y-%m-%d').date()
            #     dias = empleado_id.get_work_days_data(Datetime.from_string(fecha_inicio.strftime('%Y-%m-%d')), Datetime.from_string(fecha_hasta), calendar=empleado_id.contract_id.resource_calendar_id)
        # return dias['days']
        return diferencia_meses.days + 1

    @api.multi
    def compute_sheet(self):
        res =  super(HrPayslip, self).compute_sheet()
        for nomina in self:
            mes_nomina = int(datetime.datetime.strptime(str(nomina.date_from), '%Y-%m-%d').date().strftime('%m'))
            dia_nomina = int(datetime.datetime.strptime(str(nomina.date_to), '%Y-%m-%d').date().strftime('%d'))
            anio_nomina = int(datetime.datetime.strptime(str(nomina.date_from), '%Y-%m-%d').date().strftime('%Y'))
            valor_pago = 0
            porcentaje_pagar = 0
            for entrada in nomina.input_line_ids:
                for prestamo in nomina.employee_id.prestamo_ids:
                    anio_prestamo = int(datetime.datetime.strptime(str(prestamo.fecha_inicio), '%Y-%m-%d').date().strftime('%Y'))
                    if (prestamo.codigo == entrada.code) and ((prestamo.estado == 'nuevo') or (prestamo.estado == 'proceso')):
                        lista = []
                        for lineas in prestamo.prestamo_ids:
                            if mes_nomina == lineas.mes and anio_nomina == lineas.anio:
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
                    prestamo._compute_prestamo()

        return res

    # SALARIO PROMEDIO POR 12 MESES LABORADOS O MENOS
    def salario_promedio(self,empleado_id,fecha_final_nomina):
        historial_salario = []
        salario_meses = {}
        salario_total = 0
        salario_promedio_total = 0
        extra_ordinario_total = 0
        salario_completo = {}
        salario_sumatoria = 0
        if empleado_id.contract_ids[0].historial_salario_ids:
            posicion_historial = 0
            for linea in empleado_id.contract_ids[0].historial_salario_ids:
                if (posicion_historial+1) < len(empleado_id.contract_ids[0].historial_salario_ids):
                    historial_salario.append({'salario': linea.salario, 'fecha':linea.fecha})
                    contador_mes_historial = 0

                    llave_salario = '01-'+str(linea.fecha.month)+'-'+str(linea.fecha.year)
                    llave_salario_fecha = datetime.datetime.strptime(str(llave_salario),'%d-%m-%Y').date()
                    llave_salario_fecha_str = '01-'+str(llave_salario_fecha.month)+'-'+str(llave_salario_fecha.year)
                    # index = empleado_id.contract_ids[0].historial_salario_ids.index(linea)
                    while llave_salario_fecha < empleado_id.contract_ids[0].historial_salario_ids[posicion_historial+1].fecha:
                        salario_completo[str(llave_salario_fecha_str)] = linea.salario
                        mes = relativedelta(months=1)
                        llave_salario_fecha = llave_salario_fecha + mes
                        llave_salario_fecha_str = '01-'+str(llave_salario_fecha.month)+'-'+str(llave_salario_fecha.year)
                        contador_mes_historial += 1

                    posicion_historial += 1

            # historial_salario_ordenado = sorted(historial_salario, key=lambda k: k['fecha'],reverse=True)
            fecha_inicio_contrato = datetime.datetime.strptime(str(empleado_id.contract_ids[0].date_start),"%Y-%m-%d")
            fecha_final_contrato = datetime.datetime.strptime(str(fecha_final_nomina),"%Y-%m-%d")
            meses_laborados = (fecha_final_contrato.year - fecha_inicio_contrato.year) * 12 + (fecha_final_contrato.month - fecha_inicio_contrato.month)

            contador_mes = 0
            if meses_laborados >= 12:
                while contador_mes < 12:
                    mes = relativedelta(months=contador_mes)
                    resta_mes = fecha_final_contrato - mes
                    mes_letras = a_letras.mes_a_letras(resta_mes.month-1)
                    llave = '01-'+str(resta_mes.month)+'-'+str(resta_mes.year)
                    # llave_fecha = datetime.datetime.strptime(str(llave),'%Y-%m-%d')
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
            salario_promedio_total = empleado_id.contract_ids[0].wage
        return salario_promedio_total

    def get_inputs(self, contracts, date_from, date_to):
        res = super(HrPayslip, self).get_inputs(contracts, date_from, date_to)
        for contract in contracts:
            mes_nomina = int(datetime.datetime.strptime(str(date_from), '%Y-%m-%d').date().strftime('%m'))
            anio_nomina = int(datetime.datetime.strptime(str(date_from), '%Y-%m-%d').date().strftime('%Y'))
            dia_nomina = int(datetime.datetime.strptime(str(date_to), '%Y-%m-%d').date().strftime('%d'))
            monto_prestamo = 0
            for prestamo in contract.employee_id.prestamo_ids:
                for r in res:
                    anio_prestamo = int(datetime.datetime.strptime(str(prestamo.fecha_inicio), '%Y-%m-%d').date().strftime('%Y'))
                    if (prestamo.codigo == r['code']) and ((prestamo.estado == 'nuevo') or (prestamo.estado == 'proceso')):
                        for lineas in prestamo.prestamo_ids:
                            if mes_nomina == lineas.mes and anio_nomina == lineas.anio:
                                if self:
                                    r['amount'] = lineas.monto*(self.porcentaje_prestamo/100)
                                else:
                                    active_id = self.env.context.get('active_id')
                                    if active_id:
                                        [data] = self.env['hr.payslip.run'].browse(active_id).read(['porcentaje_prestamo'])
                                        r['amount'] = lineas.monto*(data.get('porcentaje_prestamo')/100)
            salario = self.salario_promedio(contract.employee_id,date_to)
            res.append({'name': 'Salario promedio', 'code': 'SalarioPromedio','amount': salario,'contract_id': contract.id})
            dias = self.dias_trabajados_ultimos_meses(contract.employee_id,date_from,date_to)
            res.append({'name': 'Dias Trabajados 12 Meses','code':'DiasTrabajados12Meses','amount': dias,'contract_id': contract.id})
            dias_calendario = monthrange(date_to.year, date_to.month)[1]
            res.append({'name': 'Días calendario','code':'DiasCalendario','amount': dias_calendario,'contract_id': contract.id})

        return res

    @api.onchange('employee_id', 'date_from', 'date_to','porcentaje_prestamo')
    def onchange_employee(self):
        res = super(HrPayslip, self).onchange_employee()
        mes_nomina = int(datetime.datetime.strptime(str(self.date_from), '%Y-%m-%d').date().strftime('%m'))
        anio_nomina = int(datetime.datetime.strptime(str(self.date_from), '%Y-%m-%d').date().strftime('%Y'))
        dia_nomina = int(datetime.datetime.strptime(str(self.date_to), '%Y-%m-%d').date().strftime('%d'))
        for prestamo in self.employee_id.prestamo_ids:
            anio_prestamo = int(datetime.datetime.strptime(str(prestamo.fecha_inicio), '%Y-%m-%d').date().strftime('%Y'))
            for inmput in self.input_line_ids:
                if (prestamo.codigo == inmput.code) and ((prestamo.estado == 'nuevo') or (prestamo.estado == 'proceso')):
                    for lineas in prestamo.prestamo_ids:
                        if mes_nomina == lineas.mes and anio_nomina == lineas.anio:
                            inmput.amount = lineas.monto*(self.porcentaje_prestamo/100)
        return res


    @api.model
    def get_worked_day_lines(self, contracts, date_from, date_to):
        res = super(HrPayslip, self).get_worked_day_lines(contracts,date_from,date_to)
        tipos_ausencias_ids = []

        if self.employee_id.contract_id:
            contracts = self.employee_id.contract_id
        if version_info[0] == 12:
            tipos_ausencias_ids = self.env['hr.leave.type'].search([])
        else:
            tipos_ausencias_ids = self.env['hr.holidays.status'].search([])
        ausencias_restar = []
        dias_ausentados_restar = 0
        for ausencia in tipos_ausencias_ids:
            if ausencia.descontar_nomina:
                ausencias_restar.append(ausencia.name)
        for dias in res:
            if dias['code'] in ausencias_restar:
                dias_ausentados_restar += dias['number_of_days']
        if contracts.date_start and date_from <= contracts.date_start <= date_to:
            dias_laborados = self.employee_id.get_work_days_data(Datetime.from_string(contracts.date_start), Datetime.from_string(date_to), calendar=contracts.resource_calendar_id)
            dia_inicio_contrato = int(datetime.datetime.strptime(str(contracts.date_start), '%Y-%m-%d').date().strftime('%d'))

            if version_info[0] == 12:
                if contracts.schedule_pay == 'monthly':
                    dias_laborados = 30
                if contracts.schedule_pay == 'bi-monthly':
                    dias_laborados = 15
                dias_laborados = dias_laborados - (contracts.date_start - date_from).days
                res.append({'name': 'Dias trabajados', 'sequence': 10,'code': 'TRABAJO100', 'number_of_days': (dias_laborados - dias_ausentados_restar), 'contract_id': contracts.id})
                res.append({'name': 'Dias trabajados mes', 'sequence': 10,'code': 'TRABAJOMES', 'number_of_days': (30- dia_inicio_contrato - dias_ausentados_restar), 'contract_id': contracts.id})
            else:
                res.append({'name': 'Dias trabajados', 'sequence': 10,'code': 'TRABAJO100', 'number_of_days': (dias_laborados['days'] - dias_ausentados_restar) if (dias_laborados['days'] - dias_ausentados_restar) <= 30 else 30, 'contract_id': contracts.id})
                res.append({'name': 'Dias trabajados mes', 'sequence': 10,'code': 'TRABAJOMES', 'number_of_days': (30- dia_inicio_contrato - dias_ausentados_restar), 'contract_id': contracts.id})
        elif contracts.date_end and date_from <= contracts.date_end <= date_to:
            dias_laborados = self.employee_id.get_work_days_data(Datetime.from_string(date_from), Datetime.from_string(contracts.date_end), calendar=contracts.resource_calendar_id)
            dias_trabajo = int(datetime.datetime.strptime(str(contracts.date_end), '%Y-%m-%d').date().strftime('%d'))
            res.append({'name': 'Dias trabajados', 'sequence': 10,'code': 'TRABAJO100', 'number_of_days': (dias_laborados['days'] + 1 - dias_ausentados_restar) if (dias_laborados['days'] + 1 - dias_ausentados_restar) <= 30 else 30, 'contract_id': contracts.id})
            res.append({'name': 'Dias trabajados mes', 'sequence': 10,'code': 'TRABAJOMES', 'number_of_days': (dias_trabajo - dias_ausentados_restar), 'contract_id': contracts.id})
        else:
            if contracts.schedule_pay == 'monthly':
                res.append({'name': 'Dias trabajados','sequence': 10,'code': 'TRABAJO100','number_of_days': 30 - dias_ausentados_restar, 'contract_id': contracts.id})
                res.append({'name': 'Dias trabajados mes', 'sequence': 10,'code': 'TRABAJOMES', 'number_of_days': (30 - dias_ausentados_restar), 'contract_id': contracts.id})
            if contracts.schedule_pay == 'bi-monthly':
                res.append({'name': 'Dias trabajados','sequence': 10,'code': 'TRABAJO100','number_of_days': 15 - dias_ausentados_restar, 'contract_id': contracts.id})
                res.append({'name': 'Dias trabajados mes', 'sequence': 10,'code': 'TRABAJOMES', 'number_of_days': (30 - dias_ausentados_restar), 'contract_id': contracts.id})
        return res

class HrPayslipRun(models.Model):
    _inherit = 'hr.payslip.run'

    porcentaje_prestamo = fields.Float('Prestamo (%)')
    estructura_id = fields.Many2one('hr.payroll.structure','Estructura')
    formulario = fields.Char('Número de formulario')
    total_devolucion_isr = fields.Float('Total devolución ISR')

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
                    #pago_id.post()
        return True

    @api.multi
    def close_payslip_run(self):
        for slip in self.slip_ids:
            if slip.state == 'draft':
                slip.action_payslip_done()

        res = super(HrPayslipRun, self).close_payslip_run()
        return res
