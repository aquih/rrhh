import logging
from odoo.upgrade import util

_logger = logging.getLogger(__name__)


def migrate(cr, version):
    util.records.remove_record(cr, xml_id="rrhh.act_rrhh_prestamo")
    util.records.remove_record(cr, xml_id="rrhh.act_rrhh_prestamo_2")
    util.records.remove_record(cr, xml_id="rrhh.action_rrhh_prestamo")
    util.records.remove_record(cr, xml_id="rrhh.rrhh_menu_prestamo")
    util.records.remove_view(cr, xml_id="rrhh.rrhh_view_employee_prestamo_form")
    util.records.remove_view(cr, xml_id="rrhh.rrhh_view_employee_prestamo_linea_form")
    _logger.info("Borrar vistas viejas")
