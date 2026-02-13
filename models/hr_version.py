from odoo import models, fields, api
import logging

class HrVersion(models.Model):
    _inherit = "hr.version"

    base_extra = fields.Monetary('Base Extra', digits=(16,2), tracking=True)
    bonificacion_decreto = fields.Monetary('Bonificación decreto', digits=(16,2), tracking=True)
    fecha_reinicio_labores = fields.Date('Fecha de reinicio labores')
    temporalidad_contrato = fields.Char('Temporalidad del contrato')
    calcula_indemnizacion = fields.Boolean('Calcula indemnización')
    historial_salario_ids = fields.One2many('rrhh.historial_salario', 'contrato_id', string='Historial de salario')

    # motivo_terminacion = fields.Selection([('reuncia', 'Renuncia'), ('despido', 'Despido'), ('despido_justificado', 'Despido Justificado')], 'Motivo de terminacion') no parece usarse
