import logging
from odoo.upgrade import util

_logger = logging.getLogger(__name__)


def migrate(cr, version):
    util.rename_model(cr, 'res.company.centro_trabajo', 'rrhh.centro_trabajo')
    util.rename_model(cr, 'res.company.tipo_planilla', 'rrhh.tipo_planilla')
    util.rename_model(cr, 'res.company.liquidacion', 'rrhh.liquidacion_tipo_planilla')
    _logger.info("Cambiar nombre de modelos")
