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
    util.records.remove_view(cr, xml_id="rrhh.rrhh_edit_holiday_status_form")
    util.records.remove_view(cr, xml_id="rrhh.rrhh_view_hr_payslip_form")
    util.records.remove_view(cr, xml_id="rrhh.rrhh_hr_work_entry_type_view_form")
    util.records.remove_view(cr, xml_id="rrhh.view_planilla_list")
    util.records.remove_view(cr, xml_id="rrhh.view_planilla_form")
    _logger.info("Borrar vistas viejas")
