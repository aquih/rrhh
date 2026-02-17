import logging
from odoo.upgrade import util

_logger = logging.getLogger(__name__)


def migrate(cr, version):
    util.rename_field(cr, 'rrhh.recibo', 'entrada_id', 'linea_entrada_id')
    util.rename_model(cr, 'rrhh.recibo.linea', 'rrhh.recibo.regla')
    util.rename_model(cr, 'rrhh.entrada.linea', 'rrhh.recibo.entrada')
    _logger.info("Cambiar nombre de modelos y campos de recibos")
