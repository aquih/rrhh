from odoo import api, fields, models, _
from odoo.exceptions import ValidationError
from odoo.exceptions import UserError, AccessError

class rrhh_historial_salario(models.Model):
    _name = "rrhh.historial_salario"
    _description = "Historial de salarios de un empleado."
    _order = "fecha asc"

    salario = fields.Monetary('Salario', required=True)
    anio = fields.Integer('Año', required=True)
    mes = fields.Integer('Mes', required=True)
    contrato_id = fields.Many2one('hr.version', 'Empleado')
    company_id = fields.Many2one('res.company', default=lambda self: self.env.company)
    currency_id = fields.Many2one(string="Currency", related='company_id.currency_id', readonly=True)

    # TODO: Quitar en la siguiente versión
    fecha = fields.Date('Fecha') # Ya no se usa, ahora se usa anio y mes, por qué no se pueden ingresar más de un valor por mes