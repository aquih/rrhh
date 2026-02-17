from odoo import api, fields, models, _
from odoo.exceptions import ValidationError
from odoo.exceptions import UserError, AccessError

class rrhh_historial_salario(models.Model):
    _name = "rrhh.historial_salario"
    _order = "fecha asc"

    salario = fields.Monetary('Salario', required=True)
    fecha = fields.Date('Fecha', required=True)
    contrato_id = fields.Many2one('hr.version', 'Empleado')
    company_id = fields.Many2one('res.company', default=lambda self: self.env.company, tracking=True)
    currency_id = fields.Many2one(string="Currency", related='company_id.currency_id', readonly=True)
