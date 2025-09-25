# -*- coding: utf-8 -*-
# Part of Odoo. See LICENSE file for full copyright and licensing details.

from odoo import api, fields, models, _
from odoo.exceptions import UserError
import base64
import io
import xlsxwriter

class rrhh_informe_isr(models.TransientModel):
    _name = 'rrhh.informe_isr'

    name = fields.Char('Nombre archivo')
    archivo = fields.Binary('Archivo')

    def obtener_lote(self, id):
        lote_id = self.env["hr.payslip.run"].search([("id","in",id)])
        return lote_id

    def _obtener_fecha_alta(self, nomina):
        mes = nomina.date_to.month
        anio = nomina.date_to.year
        fecha_alta = ""
        if nomina.contract_id.date_start:
            mes_contrato_inicio = nomina.contract_id.date_start.month
            anio_contrato_inicio = nomina.contract_id.date_start.year
            if anio == anio_contrato_inicio and mes == mes_contrato_inicio:
                fecha_alta = str(nomina.contract_id.date_start)
        return fecha_alta

    def _obtener_fecha_baja(self, nomina):
        mes = nomina.date_to.month
        anio = nomina.date_to.year
        fecha_baja = ""
        if nomina.contract_id.date_end:
            mes_contrato_fin = nomina.contract_id.date_end.month
            anio_contrato_fin = nomina.contract_id.date_end.year
            if anio == anio_contrato_fin and mes == mes_contrato_fin:
                fecha_baja = str(nomina.contract_id.date_end)
        return fecha_baja

    def print_report_excel(self):
        for w in self:
            dic = {}
            f = io.BytesIO()
            libro = xlsxwriter.Workbook(f)
            
            bold = libro.add_format({'bold': True})
            center = libro.add_format({'align': 'center'})
            text_wrap_bold = libro.add_format({'text_wrap': 'true', 'bold': True, 'align': 'center', 'valign': 'vcenter'})
            center_border_top = libro.add_format({'top': 1, 'align': 'center'})
            num_format = libro.add_format({'num_format': 'Q#,##0.00'})
            num_format_top = libro.add_format({'num_format': 'Q#,##0.00', 'top': 1,})
            date_format = libro.add_format({'num_format': 'dd/mm/yy', 'align': 'center'})
            
            ids = self.env.context.get('active_ids', [])
            lote_id = self.obtener_lote(ids)
            company = lote_id.company_id.name
            hoja = libro.add_worksheet("ISR")
            
            hoja.write(1, 0, lote_id.company_id.name, bold)
            hoja.write(2, 0, "Reporte general ISR", bold)
            hoja.write(3, 0, "Periodo", bold)
            hoja.write(3, 1, str(lote_id.date_start), bold)
            hoja.write(3, 2, "al", bold)
            hoja.write(3, 3, str(lote_id.date_end), bold)
            
            hoja.write(5, 0, "Correlativo", bold)
            hoja.write(5, 1, "Código", bold)
            hoja.write(5, 2, "Fecha alta", bold)
            hoja.write(5, 3, "Fecha baja", bold)
            hoja.write(5, 4, "NIT", bold)
            hoja.write(5, 5, "Nombre", bold)
            hoja.write(5, 6, "Sueldos", bold)
            hoja.write(5, 7, "Horas extras", bold)
            hoja.write(5, 8, "Bonificacion2", bold)
            hoja.write(5, 9, "Aguinaldo", bold)
            hoja.write(5, 10, "Bono 14", bold)
            hoja.write(5, 11, "Bono productividad", bold)
            hoja.write(5, 12, "Otors ingresos afectos", bold)
            hoja.write(5, 13, "Ingresos", bold)
            hoja.write(5, 14, "Aguinaldo", bold)
            hoja.write(5, 15, "Bono 14", bold)
            hoja.write(5, 16, "Cuota IGSS", bold)
            hoja.write(5, 17, "Deducciones", bold)
            hoja.write(5, 18, "Deduccion fija", bold)
            hoja.write(5, 19, "Renta imponible", bold)
            hoja.write(5, 20, "5%", bold)
            hoja.write(5, 21, "7%", bold)
            hoja.write(5, 22, "Retencion acumulada", bold)
            hoja.write(5, 23, "Retencion Descontar", bold)
            hoja.write(5, 24, "ISR descontar", bold)
            fila = 6
            
            if lote_id.slip_ids:
                correlativo = 1
                for nomina in lote_id.slip_ids:
                    fecha_alta = self._obtener_fecha_alta(nomina)
                    fecha_baja = self._obtener_fecha_baja(nomina)
                            
                    calculo_isr = self.env["hr.payslip"].calculo_isr(nomina)
                    hoja.write(fila, 0, correlativo)
                    hoja.write(fila, 1, nomina.employee_id.codigo_empleado)
                    hoja.write(fila, 2, fecha_alta)
                    hoja.write(fila, 3, fecha_baja)
                    hoja.write(fila, 4, nomina.employee_id.nit)
                    hoja.write(fila, 5, nomina.employee_id.name)
                    hoja.write(fila, 6, calculo_isr["sueldos"])
                    hoja.write(fila, 7, calculo_isr["horas_extras"])
                    hoja.write(fila, 8, calculo_isr["bonificacion_decreto"])
                    hoja.write(fila, 9, calculo_isr["aguinaldo"])
                    hoja.write(fila, 10, calculo_isr["bonoc"])
                    hoja.write(fila, 11, calculo_isr["bono_productividad"])
                    hoja.write(fila, 12, calculo_isr["otro_ingreso_afecto"])
                    hoja.write(fila, 13, calculo_isr["rubro_ingresos"])
                    hoja.write(fila, 14, calculo_isr["aguinaldo"])
                    hoja.write(fila, 15, calculo_isr["bonoc"])
                    hoja.write(fila, 16, calculo_isr["cuota_igss"])
                    hoja.write(fila, 17, calculo_isr["rubro_deducciones"])
                    hoja.write(fila, 18, calculo_isr["deduccion_fija"])
                    hoja.write(fila, 19, calculo_isr["renta_impunible"])
                    hoja.write(fila, 20, calculo_isr["rubro_renta_cinco"])
                    hoja.write(fila, 21, calculo_isr["rubro_renta_siete"])
                    hoja.write(fila, 22, calculo_isr["rubro_retencion_anual"])
                    hoja.write(fila, 23, calculo_isr["retencion_isr_descontado"])
                    hoja.write(fila, 24, calculo_isr["isr_total"])
                    correlativo += 1
                    fila += 1
                
            libro.close()
            datos = base64.b64encode(f.getvalue())
            self.write({'archivo':datos, 'name':'informe_isr.xlsx'})

        return {
            'view_type': 'form',
            'view_mode': 'form',
            'res_model': 'rrhh.informe_isr',
            'res_id': self.id,
            'view_id': False,
            'type': 'ir.actions.act_window',
            'target': 'new',
        }