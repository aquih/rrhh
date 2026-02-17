from odoo import models, fields, api
import logging

class HrVersion(models.Model):
    _inherit = "hr.version"

    base_extra = fields.Monetary('Base Extra', tracking=True, groups="hr_payroll.group_hr_payroll_user")
    bonificacion_decreto = fields.Monetary('Bonificación decreto', tracking=True, groups="hr_payroll.group_hr_payroll_user")
    fecha_reinicio_labores = fields.Date('Fecha de reinicio labores', groups="hr_payroll.group_hr_payroll_user")
    temporalidad_contrato = fields.Char('Temporalidad del contrato', groups="hr_payroll.group_hr_payroll_user")
    calcula_indemnizacion = fields.Boolean('Calcula indemnización', groups="hr_payroll.group_hr_payroll_user")
    historial_salario_ids = fields.One2many('rrhh.historial_salario', 'contrato_id', string='Historial de salario', groups="hr_payroll.group_hr_payroll_user")

    # TODO: Quitar en la siguiente versión
    motivo_terminacion = fields.Selection([('reuncia', 'Renuncia'), ('despido', 'Despido'), ('despido_justificado', 'Despido Justificado')], 'Motivo de terminacion', groups="hr_payroll.group_hr_payroll_user") # no parece usarse
