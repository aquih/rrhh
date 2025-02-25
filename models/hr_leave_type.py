# Part of Odoo. See LICENSE file for full copyright and licensing details.


from odoo import api, fields, models


class HolidaysType(models.Model):
    _inherit = "hr.leave.type"

    suspension_igss = fields.Boolean(string="Suspensión IGSS")