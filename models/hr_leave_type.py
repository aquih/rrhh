from odoo import api, fields, models

class HolidaysType(models.Model):
    _inherit = "hr.leave.type"

    # TODO: Quitar en la siguiente versión
    suspension_igss = fields.Boolean(string="Suspensión IGSS") # usar mejor validación contra un tipo nuevo de ausencia