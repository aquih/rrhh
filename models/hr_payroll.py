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

class HrPayslip(models.Model):
    _inherit = 'hr.payslip'

    porcentaje_prestamo = fields.Float(related="payslip_run_id.porcentaje_prestamo",string='Prestamo (%)',store=True)
    etiqueta_empleado_ids = fields.Many2many('hr.employee.category',string='Etiqueta empleado', related='employee_id.category_ids')

    def action_payslip_done(self):
        res = super(HrPayslip, self).action_payslip_done()
        for slip in self:
            logging.warn('1')
            if slip.move_id:
                logging.warn('2')
                slip.move_id.button_cancel()
                for line in slip.move_id.line_ids:
                    logging.warn('3')
                    line.analytic_account_id = slip.contract_id.analytic_account_id.id
                slip.move_id.post()
        return res

    # Dias trabajdas de los ultimos 12 meses hasta la fecha
    def dias_trabajados_ultimos_meses(self,empleado_id,fecha):
        dias = {'days': 0}
        if empleado_id.contract_id.date_start:
            fecha_nomina = datetime.datetime.strptime(str(fecha), '%Y-%m-%d').date()
            fecha_contrato = datetime.datetime.strptime(str(empleado_id.contract_id.date_start), '%Y-%m-%d').date()
            diferencia_meses = relativedelta(fecha_nomina,fecha_contrato)
            empleado = self.env['hr.employee'].browse(empleado_id)
            if int(diferencia_meses.years) == 0:
                dias = empleado_id._get_work_days_data(Datetime.from_string(empleado_id.contract_id.date_start), Datetime.from_string(fecha), calendar=empleado_id.contract_id.resource_calendar_id)
            else:
                mes = relativedelta(months=12)
                fecha_inicio = datetime.datetime.strptime(str(fecha_nomina - mes), '%Y-%m-%d').date()
                dias = empleado_id._get_work_days_data(Datetime.from_string(fecha_inicio.strftime('%Y-%m-%d')), Datetime.from_string(fecha), calendar=empleado_id.contract_id.resource_calendar_id)
        return dias['days']

    def compute_sheet(self):
        res =  super(HrPayslip, self).compute_sheet()
        for nomina in self:
            if nomina.contract_id:
                entradas = self._obtener_entrada(nomina.contract_id)
                if entradas:
                    for entrada in entradas:
                        entrada_id = self.env['hr.payslip.input'].create({'payslip_id': nomina.id,'input_type_id': entrada.id})

                self.calculo_rrhh(nomina.contract_id,nomina,nomina.date_to)
            mes_nomina = int(datetime.datetime.strptime(str(nomina.date_from), '%Y-%m-%d').date().strftime('%m'))
            dia_nomina = int(datetime.datetime.strptime(str(nomina.date_to), '%Y-%m-%d').date().strftime('%d'))
            anio_nomina = int(datetime.datetime.strptime(str(nomina.date_from), '%Y-%m-%d').date().strftime('%Y'))
            valor_pago = 0
            porcentaje_pagar = 0
            for entrada in nomina.input_line_ids:
                for prestamo in nomina.employee_id.prestamo_ids:
                    anio_prestamo = int(datetime.datetime.strptime(str(prestamo.fecha_inicio), '%Y-%m-%d').date().strftime('%Y'))
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
        return res

    def _obtener_entrada(self,contrato_id):
        entradas = False
        if contrato_id.structure_type_id and contrato_id.structure_type_id.default_struct_id:
            if contrato_id.structure_type_id.default_struct_id.input_line_type_ids:
                entradas = [entrada for entrada in contrato_id.structure_type_id.default_struct_id.input_line_type_ids]
        return entradas

    def calculo_rrhh(self,contrato_id,nomina,date_to):
        salario = self.salario_promedio(contrato_id.employee_id,contrato_id.company_id.salario_ids.ids)
        dias = self.dias_trabajados_ultimos_meses(contrato_id.employee_id,date_to)
        for entrada in nomina.input_line_ids:
            if entrada.input_type_id.code == 'SalarioPromedio':
                entrada.amount = salario
            if entrada.input_type_id.code == 'DiasTrabajados12Meses':
                entrada.amount = dias
        return True

    def salario_promedio(self, empleado_id, reglas):
        fecha_hoy = datetime.datetime.now()
        salario = 0
        nomina_ids = self.env['hr.payslip'].search([['employee_id', '=', empleado_id.id]])
        nominas = []
        contador = 1
        meses_nominas = []
        while contador <= 12:
            mes = relativedelta(months=contador)
            resta_mes = fecha_hoy - mes
            for nomina in nomina_ids:
                nomina_mes = datetime.datetime.strptime(str(nomina.date_from),"%Y-%m-%d")
                if nomina_mes.month == resta_mes.month and nomina_mes.year == resta_mes.year:
                    if resta_mes not in meses_nominas:
                        meses_nominas.append({resta_mes.month: resta_mes.month})
                    else:
                        meses_nominas[resta_mes.month] = resta_mes.month
                    nominas.append(nomina)
                    for linea in nomina.line_ids:
                        if linea.salary_rule_id.id in reglas:
                            salario += linea.total
            contador += 1
        promedio = salario
        if len(meses_nominas) > 0:
            promedio = salario / len(meses_nominas)
        return promedio


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
        logging.warn('el res 2')
        datos = self.horas_sumar(res)
        for r in res:
            tipo_id = self.env['hr.work.entry.type'].search([('id','=',r['work_entry_type_id'])])
            if tipo_id and tipo_id.is_leave == False:
                r['number_of_hours'] += datos['horas']
                r['number_of_days'] += datos['dias']
        logging.warn(res)
        return res

    @api.onchange('employee_id','struct_id','contract_id', 'date_from', 'date_to','porcentaje_prestamo')
    def _onchange_employee(self):
        res = super(HrPayslip, self)._onchange_employee()
        mes_nomina = int(datetime.datetime.strptime(str(self.date_from), '%Y-%m-%d').date().strftime('%m'))
        anio_nomina = int(datetime.datetime.strptime(str(self.date_from), '%Y-%m-%d').date().strftime('%Y'))
        dia_nomina = int(datetime.datetime.strptime(str(self.date_to), '%Y-%m-%d').date().strftime('%d'))
        for prestamo in self.employee_id.prestamo_ids:
            anio_prestamo = int(datetime.datetime.strptime(str(prestamo.fecha_inicio), '%Y-%m-%d').date().strftime('%Y'))
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

    def close_payslip_run(self):
        for slip in self.slip_ids:
            if slip.state == 'draft':
                slip.action_payslip_done()

        res = super(HrPayslipRun, self).close_payslip_run()
        return res
