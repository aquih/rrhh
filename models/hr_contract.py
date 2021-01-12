# -*- coding: utf-8 -*-

from odoo import models, fields, api

class ContractType(models.Model):
    _inherit = "hr.contract.type"

    calcula_indemnizacion = fields.Boolean('Calcula indemnizacion')

class Contract(models.Model):
    _inherit = "hr.contract"

    motivo_terminacion = fields.Selection([
        ('reuncia', 'Renuncia'),
        ('despido', 'Despido'),
        ('despido_justificado', 'Despido Justificado'),
        ], 'Motivo de terminacion')
    base_extra = fields.Monetary('Base Extra', digits=(16,2), track_visibility='onchange')
    # salario_extra_ordinario_id = fields.Many2many('hr.salary.rule', 'salario_extra_regla_rel', 'contrato_id', 'regla_id', string='Salario extra ordinario')
    # igss_id = fields.Many2many('hr.salary.rule', 'igss_regla_rel', 'contrato_id', 'regla_id', string='IGSS')
    # otras_deducciones_legales_id = fields.Many2many('hr.salary.rule', 'otras_deducciones_legales_regla_rel', 'contrato_id', 'regla_id', string='Otras deducciones legales')
    # total_deducciones_id = fields.Many2many('hr.salary.rule', 'total_deducciones_regla_rel', 'contrato_id', 'regla_id', string='Total deducciones')
    # decreto_42_92_id = fields.Many2many('hr.salary.rule', 'decreto_42_92_regla_rel', 'contrato_id', 'regla_id', string='Decreto 42-92')
    # bonificacion_incentivo_id = fields.Many2many('hr.salary.rule', 'bonificacion_incentivo_regla_rel', 'contrato_id', 'regla_id', string='Bonificacion Incentivo')
    # comisiones_id = fields.Many2many('hr.salary.rule', 'comisiones_regla_rel', 'contrato_id', 'regla_id', string='Comisiones')
    # septimos_asuetos_id = fields.Many2many('hr.salary.rule', 'septimos_asuetos_regla_rel', 'contrato_id', 'regla_id', string='Septimos asuetos')
    # vacaciones_id = fields.Many2many('hr.salary.rule', 'vacaciones_regla_rel', 'contrato_id', 'regla_id', string='Vacaciones')
    # liquido_recibir_id = fields.Many2many('hr.salary.rule', 'liquido_recibir_regla_rel', 'contrato_id', 'regla_id', string='Liquido a recibir')
    wage = fields.Monetary('Wage', digits=(16, 2), required=True, help="Employee's monthly gross wage.",track_visibility='onchange')
    fecha_reinicio_labores = fields.Date('Fecha de reinicio labores')
    temporalidad_contrato = fields.Char('Teporalidad del contrato')
    calcula_indemnizacion = fields.Boolean('Calcula indemnizacion')
