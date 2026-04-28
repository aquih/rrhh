# -*- encoding: utf-8 -*-

from odoo import models, fields, api, _
import time
import base64
import xlsxwriter
import io
import logging

class rrhh_planilla_wizard(models.TransientModel):
    _name = 'rrhh.planilla.wizard'
    _description = 'Wizard de planilla'

    nomina_id = fields.Many2one('hr.payslip.run', 'Nomina', default=lambda self: self.env['hr.payslip.run'].browse(self._context.get('active_id')), required=True)
    planilla_id = fields.Many2one('rrhh.planilla', 'Planilla', required=True)
    archivo = fields.Binary('Archivo')
    name =  fields.Char('File Name', default='planilla.xlsx')
    agrupado  = fields.Boolean('Agrupado por cuenta analítica')

    def print_report(self):
        datas = {'ids': self.env.context.get('active_ids', [])}
        res = self.read([])
        res = res and res[0] or {}
        datas['form'] = res
        return self.env.ref('rrhh.action_planilla').with_context(landscape=True).report_action([], data=datas)

    def buscar_partida_nominas(self, slip_ids):
        cantidad_nominas_partida = 0
        partidas = {}
        for slip in slip_ids:
            if slip.move_id:
                if slip.id not in partidas:
                    partidas[slip.move_id.id] = 0
        return partidas

    def generar_excel(self):
        for w in self:
            dict = {}
            dict['planilla_id'] = [w.planilla_id.id, w.planilla_id.name]
            dict['nomina_id'] = [w.nomina_id.id, w.nomina_id.name]
            dict['agrupado'] = w['agrupado']
            reporte = self.env['report.rrhh.planilla_pdf'].reporte(dict)

            f = io.BytesIO()
            libro = xlsxwriter.Workbook(f)
            formato_fecha = libro.add_format({'num_format': 'dd/mm/yy'})

            if w.agrupado:
                for cuenta in reporte['cuentas_analiticas']:
                    hoja = libro.add_worksheet(cuenta)

                    hoja.write(0, 0, 'Planilla')
                    hoja.write(0, 1, w.nomina_id.name)
                    hoja.write(0, 2, 'Periodo')
                    hoja.write(0, 3, w.nomina_id.date_start, formato_fecha)
                    hoja.write(0, 4, w.nomina_id.date_end, formato_fecha)
                    
                    linea = 2
                    for puesto in reporte['puestos'][cuenta]:
    
                        num = 1

                        hoja.write(linea, 0, puesto)
                        linea += 2

                        hoja.write(linea, 0, 'No')
                        hoja.write(linea, 1, 'Cod. de empleado')
                        hoja.write(linea, 2, 'Nombre de empleado')
                        hoja.write(linea, 3, 'Fecha de ingreso')
                        hoja.write(linea, 4, 'Dias')

                        columna = 4
                        for nombre_columna in reporte['columnas']:
                            columna += 1
                            hoja.write(linea, columna, nombre_columna)
    
                        hoja.write(linea, columna+1, 'Banco a depositar')
                        hoja.write(linea, columna+2, 'Cuenta a depositar')
                        hoja.write(linea, columna+3, 'Observaciones')

                        linea += 1
                        for linea_reporte in reporte['lineas'][cuenta][puesto]['datos']:
                            hoja.write(linea, 0, linea_reporte['estatico']['numero'] or '')
                            hoja.write(linea, 1, linea_reporte['estatico']['codigo_empleado'] or '')
                            hoja.write(linea, 2, linea_reporte['estatico']['nombre_empleado'] or '')
                            hoja.write(linea, 3, linea_reporte['estatico']['fecha_ingreso'] or '', formato_fecha)
                            hoja.write(linea, 4, linea_reporte['estatico']['dias'] or '')

                            columna = 4
                            for l in linea_reporte['dinamico']:
                                columna += 1
                                hoja.write(linea, columna, l)
        
                            hoja.write(linea, columna+1, linea_reporte['estatico']['banco_depositar'] or '')
                            hoja.write(linea, columna+2, linea_reporte['estatico']['cuenta_depositar'] or '')
                            hoja.write(linea, columna+3, linea_reporte['estatico']['observaciones'] or '')
                            linea += 1

                        hoja.write(linea, 3, 'TOTALES')
                        columna = 4
                        for t in reporte['lineas'][cuenta][puesto]['totales']:
                            columna += 1
                            hoja.write(linea, columna, t)

                        linea += 1

                    linea += 2
                    columna = 4
                    for t in reporte['columnas']:
                        columna += 1
                        hoja.write(linea, columna, t)

                    linea += 1
                    hoja.write(linea, 3, 'TOTALES')
                    columna = 4
                    for t in reporte['suma'][cuenta]:
                        columna += 1
                        hoja.write(linea, columna, t)
            
            else:
                hoja = libro.add_worksheet('reporte')

                hoja.write(0, 0, 'Planilla')
                hoja.write(0, 1, w.nomina_id.name)
                hoja.write(0, 2, 'Periodo')
                hoja.write(0, 3, w.nomina_id.date_start, formato_fecha)
                hoja.write(0, 4, w.nomina_id.date_end, formato_fecha)

                linea = 2

                hoja.write(linea, 0, 'No')
                hoja.write(linea, 1, 'Cod. de empleado')
                hoja.write(linea, 2, 'Nombre de empleado')
                hoja.write(linea, 3, 'Fecha de ingreso')
                hoja.write(linea, 4, 'Puesto')
                hoja.write(linea, 5, 'Dias')

                columna = 5
                for nombre_columna in reporte['columnas']:
                    columna += 1
                    hoja.write(linea, columna, nombre_columna)

                hoja.write(linea, columna+1, 'Banco a depositar')
                hoja.write(linea, columna+2, 'Cuenta a depositar')
                hoja.write(linea, columna+3, 'Observaciones')

                linea += 1
                for empleado in reporte['no_agrupado']:
                    hoja.write(linea, 0, empleado['numero'] or '')
                    hoja.write(linea, 1, empleado['codigo_empleado'] or '')
                    hoja.write(linea, 2, empleado['nombre_empleado'] or '')
                    hoja.write(linea, 3, empleado['fecha_ingreso'] or '', formato_fecha)
                    hoja.write(linea, 4, empleado['puesto'] or '')
                    hoja.write(linea, 5, empleado['dias'] or '')

                    columna = 5
                    for l in empleado['columnas']:
                        columna += 1
                        hoja.write(linea, columna, l)

                    hoja.write(linea, columna+1, empleado['banco_depositar'] or '')
                    hoja.write(linea, columna+2, empleado['cuenta_depositar'] or '')
                    hoja.write(linea, columna+3, empleado['observaciones'] or '')
                    linea += 1

                hoja.write(linea,4, 'GRAN TOTAL')
                columna = 5
                for t in reporte['total']:
                    columna += 1
                    hoja.write(linea, columna, t)

            libro.close()
            datos = base64.b64encode(f.getvalue())
            self.write({'archivo': datos})
            return {
                'context': self.env.context,
                'view_type': 'form',
                'view_mode': 'form',
                'res_model': 'rrhh.planilla.wizard',
                'res_id': self.id,
                'view_id': False,
                'type': 'ir.actions.act_window',
                'target': 'new',
            }
