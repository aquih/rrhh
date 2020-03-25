# -*- coding: utf-8 -*-

from odoo import api, models
import datetime
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
        reporte['cuentas_analiticas'] = []
        reporte['puestos'] = {}
        reporte['lineas'] = []
        reporte['suma'] = {}
        reporte['total'] = {}
        reporte['total'] = []
        totals1 = [0 for c in planilla.columna_id]
        totals1.append(0)
        reporte['total'] = totals1
        
        columnas = []
        for columna in planilla.columna_id:
            columnas.append(columna.name)

        columnas.append('Liquido a recibir')

        lineas = {}
        numero = 1
        for slip in nomina.slip_ids:
            if slip.contract_id.analytic_account_id.name:
                llave = slip.contract_id.analytic_account_id.name
            else:
                llave = 'Indefinido'
            
            
            if llave not in lineas:
                lineas[llave] = {}
                
            if slip.employee_id.job_id.name not in lineas[llave]:
                lineas[llave][slip.employee_id.job_id.name] = {}
                lineas[llave][slip.employee_id.job_id.name]['datos'] = []

                lineas[llave][slip.employee_id.job_id.name]['totales'] = []

                totales = [0 for c in planilla.columna_id]
                totales.append(0)

                lineas[llave][slip.employee_id.job_id.name]['totales'] = totales

            if llave not in reporte['cuentas_analiticas']:
                reporte['cuentas_analiticas'].append(llave)
                reporte['suma'][llave] = []
                totals = [0 for c in planilla.columna_id]
                totals.append(0)
                reporte['suma'][llave] = totals

            if llave not in reporte['puestos']:
                reporte['puestos'][llave] = []
            
            if slip.employee_id.job_id.name not in reporte['puestos'][llave]:
                reporte['puestos'][llave].append(slip.employee_id.job_id.name)


            linea = {'estatico': {}, 'dinamico': []}
            linea['estatico']['numero'] = numero
            linea['estatico']['codigo_empleado'] = slip.employee_id.codigo_empleado
            linea['estatico']['nombre_empleado'] = slip.employee_id.name
            linea['estatico']['fecha_ingreso'] = datetime.datetime.strptime(str(slip.contract_id.date_start), "%Y-%m-%d").strftime('%d/%m/%Y')
#            linea['estatico']['fecha_ingreso'] = slip.contract_id.date_start
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
                lineas[llave][slip.employee_id.job_id.name]['totales'][x] += total_columna
                reporte['suma'][llave][x] += total_columna
                reporte['total'][x] += total_columna
                x += 1
            
            linea['dinamico'].append(total_salario)
            lineas[llave][slip.employee_id.job_id.name]['totales'][- 1] += total_salario
            reporte['suma'][llave][- 1] += total_salario
            reporte['total'][- 1] += total_salario
            
            linea['estatico']['banco_depositar'] = slip.employee_id.bank_account_id.bank_id.name
            linea['estatico']['cuenta_depositar'] = slip.employee_id.bank_account_id.acc_number
            linea['estatico']['observaciones'] = slip.note
            if slip.move_id and len(slip.move_id.line_ids) > 0 and slip.move_id.line_ids[0].analytic_account_id:
                linea['estatico']['cuenta_analitica'] = slip.move_id.line_ids[0].analytic_account_id.name
            else:
                linea['estatico']['cuenta_analitica'] = llave
            lineas[llave][slip.employee_id.job_id.name]['datos'].append(linea)
            
        
        reporte['columnas'] = columnas
        reporte['lineas'] = lineas

        logging.getLogger('reporte').warn(reporte)
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
