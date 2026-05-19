# -*- encoding: utf-8 -*-

from odoo import api, models, fields
import time
import datetime
from datetime import date
from datetime import datetime, date, time, timedelta
from odoo.fields import Date, Datetime
import logging

class ReportLibroSalarios(models.AbstractModel):
    _name = 'report.rrhh.libro_salarios'
    _description = 'Libro de salarios (para el MTPS)'

    def obtener_empleado(self, id):
        return self.env['hr.employee'].search([('id', '=', id), '|', ('active', '=', True), ('active', '=', False)])

    # Usar mismo algoritmo que en planilla
    def obtener_dias(self, payslip):
        dias = 0
        work = -1
        trabajo = -1
        for d in payslip.worked_days_line_ids:
            if d.code == 'TRABAJO100':
                trabajo = d.number_of_days
            elif d.code == 'WORK100':
                work = d.number_of_days
        if trabajo >= 0:
            dias += trabajo
        else:
            dias += work

        return dias

    def domingos_trabajados(self, fecha_inicio, fecha_fin):
        cantidad_dias = (fecha_inicio - fecha_fin).days
        cantidad_domingos = 0

        for incr in range(cantidad_dias):
            dia = fecha_inicio + timedelta(days=incr)
            if dia.weekday() == 6:
                cantidad_domingos += 1

        return cantidad_domingos

    def obtener_payslips(self, id, anio):
        payslips = self.env['hr.payslip'].search([['employee_id', '=', id]], order="date_to asc")
        nominas_lista = []
        numero_orden = 0
        for payslip in payslips:
            nomina_anio = payslip.date_to.year
            if anio == nomina_anio:
                salario = 0
                dias_trabajados = self.obtener_dias(payslip)
                ordinarias = 0
                extra_ordinarias = 0
                ordinario = 0
                extra_ordinario = 0
                igss = 0
                isr = 0
                anticipos = 0
                bonificacion = 0
                bono = 0
                aguinaldo = 0
                indemnizacion = 0
                septimos_asuetos = 0
                vacaciones = 0
                decreto = 0
                fija = 0
                variable = 0
                otras_deducciones = 0
                otros_salarios = 0
                boni_incentivo_decreto = 0
                dev_isr_otro = 0
                work = -1
                trabajo = -1
                dias_calculados = self.obtener_dias(payslip)

                for linea in payslip.line_ids:
                    if linea.salary_rule_id.id in payslip.company_id.salario_ids.ids:
                        salario += linea.total
                    if linea.salary_rule_id.id in payslip.company_id.extras_ordinarias_ids.ids:
                        extra_ordinarias = sum([entrada.amount for entrada in payslip.input_line_ids if linea.code == entrada.code])
                    if linea.salary_rule_id.id in payslip.company_id.ordinario_ids.ids:
                        ordinario += linea.total
                    if linea.salary_rule_id.id in payslip.company_id.extra_ordinario_ids.ids:
                        extra_ordinario += linea.total
                    if linea.salary_rule_id.id in payslip.company_id.igss_ids.ids:
                        igss += linea.total
                    if linea.salary_rule_id.id in payslip.company_id.isr_ids.ids:
                        isr += linea.total
                    if linea.salary_rule_id.id in payslip.company_id.anticipos_ids.ids:
                        anticipos += linea.total
                    if linea.salary_rule_id.id in payslip.company_id.bonificacion_ids.ids:
                        bonificacion += linea.total
                    if linea.salary_rule_id.id in payslip.company_id.bono_ids.ids:
                        bono += linea.total
                    if linea.salary_rule_id.id in payslip.company_id.aguinaldo_ids.ids:
                        aguinaldo += linea.total
                    if linea.salary_rule_id.id in payslip.company_id.indemnizacion_ids.ids:
                        indemnizacion += linea.total
                    if linea.salary_rule_id.id in payslip.company_id.vacaciones_ids.ids:
                        vacaciones += linea.total
                    if linea.salary_rule_id.id in payslip.company_id.decreto_ids.ids:
                        decreto += linea.total
                    if linea.salary_rule_id.id in payslip.company_id.fija_ids.ids:
                        fija += linea.total
                    if linea.salary_rule_id.id in payslip.company_id.variable_ids.ids:
                        variable += linea.total
                    if linea.salary_rule_id.id in payslip.company_id.otro_salario_ids.ids:
                        otros_salarios += linea.total
                    if linea.salary_rule_id.id in payslip.company_id.boni_incentivo_decreto_ids.ids:
                        boni_incentivo_decreto += linea.total
                    if linea.salary_rule_id.id in payslip.company_id.devolucion_isr_otro_ids.ids:
                        dev_isr_otro += linea.total

                ordinarias = dias_trabajados * 8 if dias_trabajados <= 31 else 0
                domingos = self.domingos_trabajados(payslip.date_from, payslip.date_to)
                sueldo_diario = 0
                if dias_trabajados > 0:
                    sueldo_diario = salario / 30

                septimos_asuetos = sueldo_diario * domingos
                ordinario_final = ordinario - septimos_asuetos
                ordinario = ordinario_final
                otras_deducciones = anticipos
                total_deducciones = igss + otras_deducciones + isr
                bono_agui_indem = bono + aguinaldo + indemnizacion
                numero_orden += 1
                total_salario_devengado =  ordinario + extra_ordinario + septimos_asuetos + vacaciones + otros_salarios
                nominas_lista.append({
                    'orden': numero_orden,
                    'fecha_inicio': payslip.date_from,
                    'fecha_fin': payslip.date_to,
                    'moneda_id': payslip.company_id.currency_id,
                    'salario': salario,
                    'dias_trabajados': dias_trabajados,
                    'ordinarias': ordinarias,
                    'extra_ordinarias': extra_ordinarias,
                    'ordinario': ordinario,
                    'extra_ordinario': extra_ordinario,
                    'septimos_asuetos': septimos_asuetos,
                    'vacaciones': vacaciones,
                    'total_salario_devengado': total_salario_devengado,
                    'igss': igss,
                    'isr': isr,
                    'anticipos': anticipos,
                    'otras_deducciones': otras_deducciones,
                    'total_deducciones': total_deducciones,
                    'bonificacion_id': bonificacion,
                    'boni_incentivo_decreto': boni_incentivo_decreto,
                    'variable': variable,
                    'dev_isr_otro': dev_isr_otro,
                    'bono_agui_indem': bono_agui_indem,
                    'otros_salarios': otros_salarios,
                    'liquido_recibir': total_salario_devengado + total_deducciones +bono_agui_indem+ boni_incentivo_decreto + dev_isr_otro
                })
                
        return nominas_lista

    @api.model
    def _get_report_values(self, docids, data=None):
        data = data if data is not None else {}
        model = 'rrhh.libro_salarios'
        docs = data.get('ids', data.get('active_ids'))
        anio = data.get('form', {}).get('anio', False)
        folio_inicial = data.get('form', {}).get('folio_inicial', False)
        return {
            'doc_ids': docids,
            'doc_model': model,
            'docs': docs,
            'anio': anio,
            'folio_inicial': folio_inicial,
            'empleado': self.obtener_empleado,
            'payslips': self.obtener_payslips,
            'current_company_id': self.env.company,
        }
