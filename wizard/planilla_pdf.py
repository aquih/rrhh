# -*- coding: utf-8 -*-

from odoo import api, models
import logging

class report_planilla_pdf(models.AbstractModel):
    _name = 'report.rrhh.planilla_pdf'

    
    
    def reporte(self, datos):
        logging.getLogger('datos...').warn(datos)
        planilla = self.env['rrhh.planilla'].browse(datos['planilla_id'][0])
        nomina = self.env['hr.payslip.run'].browse(datos['nomina_id'][0])
        logging.getLogger('nomina.name').warn(nomina.name)
        reporte = {}
        reporte['encabezado'] = {} 
        reporte['encabezado']['nomina'] = nomina.name
        reporte['cuentas_analiticas'] = {}
        reporte['lineas'] = []

        columnas = []
        for columna in planilla.columna_id:
            columnas.append(columna.name)

        columnas.append('Liquido a recibir')

        lineas = {}
        numero = 1
        for slip in nomina.slip_ids:
            if slip.contract_id.analytic_account_id.name:
                llave = slip.contract_id.analytic_account_id.name
                if llave not in lineas:
                    lineas[slip.contract_id.analytic_account_id.name] = {}
                    lineas[slip.contract_id.analytic_account_id.name]['datos'] = []
                    reporte['cuentas_analiticas'][slip.contract_id.analytic_account_id.id] = slip.contract_id.analytic_account_id.name
            else:
                llave = 'Indefinido'
                if llave not in lineas:
                    lineas['Indefinido'] = {}
                    lineas['Indefinido']['datos'] = []
                    reporte['cuentas_analiticas'][0] = 'Indefinido'

            linea = {'estatico': {}, 'dinamico': []}
            linea['estatico']['numero'] = numero
            linea['estatico']['codigo_empleado'] = slip.employee_id.codigo_empleado
            linea['estatico']['nombre_empleado'] = slip.employee_id.name
            linea['estatico']['fecha_ingreso'] = slip.contract_id.date_start
            linea['estatico']['puesto'] = slip.employee_id.job_id.name

            dias = 0
            work = -1
            trabajo = -1
            for d in slip.worked_days_line_ids:
                if d.code == 'TRABAJO100':
                    trabajo = d.number_of_days
                elif d.code == 'WORK100':
                    work = d.number_of_days
            if trabajo >= 0:
                dias += trabajo
            else:
                dias += work
            linea['estatico']['dias'] = dias

            totales = [0 for c in planilla.columna_id]
            totales.append(0)

            total_salario = 0
            x = 0
            for c in planilla.columna_id:
                reglas = [x.id for x in c.regla_id]
                entradas = [x.name for x in c.entrada_id]
                total_columna = 0
                for r in slip.line_ids:
                    if r.salary_rule_id.id in reglas:
                        total_columna += r.total
                for r in slip.input_line_ids:
                    if r.name in entradas:
                        total_columna += r.amount
                if c.sumar:
                    total_salario += total_columna

                linea['dinamico'].append(total_columna)
                totales[x] += total_columna
                x += 1

            linea['dinamico'].append(total_salario)
            totales[len(totales) - 1] += total_salario
            linea['estatico']['banco_depositar'] = slip.employee_id.bank_account_id.bank_id.name
            linea['estatico']['cuenta_depositar'] = slip.employee_id.bank_account_id.acc_number
            linea['estatico']['observaciones'] = slip.note
            if slip.move_id and len(slip.move_id.line_ids) > 0 and slip.move_id.line_ids[0].analytic_account_id:
                linea['estatico']['cuenta_analitica'] = slip.move_id.line_ids[0].analytic_account_id.name.name
            else:
                linea['estatico']['cuenta_analitica'] = llave
            lineas[llave]['datos'].append(linea)

        lineas[llave]['totales'] = []
        lineas[llave]['totales'].append(totales)            
        reporte['columnas'] = columnas
        reporte['lineas'] = lineas
        
        logging.warn(reporte)
        return reporte

    @api.model
    def get_report_values(self, docids, data=None):
        model = self.env.context.get('active_model')
        docs = self.env[model].browse(self.env.context.get('active_ids', []))

        return {
            'doc_ids': self.ids,
            'doc_model': model,
            'data': data['form'],
            'docs': docs,
            'reporte': self.reporte,
        }
