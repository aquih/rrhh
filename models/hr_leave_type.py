from odoo import api, fields, models

class HolidaysType(models.Model):
    _inherit = "hr.leave.type"

    # TODO: Creo que se debería quitar
    suspension_igss = fields.Boolean(string="Suspensión IGSS")