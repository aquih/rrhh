# -*- coding: utf-8 -*-
from odoo import api, fields, models
from odoo.release import version_info

if version_info[0] == 12:
    class HolidaysType(models.Model):
        _inherit = "hr.leave.type"

        descontar_nomina = fields.Boolean('Descontar en nómina')
        codigo = fields.Char('codigo')
else:
    class HolidaysStatus(models.Model):
        _inherit = "hr.holidays.status"

        descontar_nomina = fields.Boolean('Descontar en nómina')
        codigo = fields.Char('codigo')
