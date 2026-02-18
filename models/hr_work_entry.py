# -*- coding: utf-8 -*-
from odoo import api, fields, models
from odoo.release import version_info

class HrWorkEntryType(models.Model):
    _inherit = "hr.work.entry.type"

    # TODO: Quitar en la siguiente versión
    descontar_nomina = fields.Boolean('Descontar en nómina') # usar mejor la forma nativa de calcular horas y días de Odoo