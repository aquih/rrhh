import logging

_logger = logging.getLogger(__name__)


def migrate(cr, version):
    cr.execute("UPDATE rrhh_historial_salario SET anio = EXTRACT(year FROM fecha)")
    cr.execute("UPDATE rrhh_historial_salario SET mes = EXTRACT(month FROM fecha)")
    _logger.info("Actualizar valores")
