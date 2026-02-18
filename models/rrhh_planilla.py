from odoo import models, fields, api

class rrhh_planilla(models.Model):
    _name = 'rrhh.planilla'
    _description = 'Estructura de planilla'

    name = fields.Char('Nombre', required=True)
    descripcion = fields.Char('Descripción')
    columna_id = fields.One2many('rrhh.planilla.columna', 'planilla_id', 'Columnas')

class rrhh_planilla_columna(models.Model):
    _name = 'rrhh.planilla.columna'
    _description = 'Estructura de columna de planilla'
    _order = 'sequence, name'

    planilla_id = fields.Many2one('rrhh.planilla', 'Planilla', required=False)
    name = fields.Char('Nombre', required=True)
    sequence = fields.Integer('Secuencia', required=True, index=True, default=5)
    regla_id = fields.Many2many('hr.salary.rule', string='Reglas')
    entrada_id = fields.Many2many('hr.payslip.input.type', string='Entradas')
    sumar = fields.Boolean('Sumar en liquido a recibir', help="Seleccionar si se desea que se tome en cuenta en la suma del liquido a recibir.")