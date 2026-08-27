from odoo import models, fields, api

class res_company(models.Model):
    _inherit = 'res.company'

    version_mensaje = fields.Char('Version del mensaje')
    numero_patronal = fields.Char('Numero patronal')
    tipo_planilla = fields.Selection([('0', 'Producción'), ('1', 'Pruebas')], 'Produccióno o pruebas')
    representante_legal_id = fields.Many2one('hr.employee', 'Representante legal')
    barrio_colonia = fields.Char('Barrio o Colonia')
    zona = fields.Char('Zona donde se ubica')
    nomenclatura = fields.Char('Nomenclatura')
    sindicato = fields.Selection([('si', 'Si'), ('no', 'No')], 'Existe sindicato')
    contratar_personal = fields.Selection([('si', 'Si'), ('no', 'No')], 'Contratar nuevo personal')
    contabilidad_completa = fields.Selection([('si', 'Si'), ('no', 'No')], 'Contabilidad completa')
    jefe_recursos_humanos_id = fields.Many2one('hr.employee', 'Jefe de recursos humanos')
    anio_inicio_operaciones = fields.Integer('Año inicio de operaciones')
    tamanio_empresa_ventas = fields.Char('Tamaño de de empresa por ventas anuales')
    tamanio_empresa_trabajadores = fields.Char('Tamaño de empresa según cantidad de Trabajadores')
    actividad_gran_grupo = fields.Char('Actividad gran grupo')
    sub_actividad_economica = fields.Char('Sub actividad económica')
    ocupacion_grupo = fields.Char('Ocupación grupo')
    monto_deduccion_fija = fields.Float("Deducción fija")
    isr_sueldo_base_extra = fields.Boolean('ISR sueldo base extra')
    proyectar_bonificaciones_adicionales = fields.Boolean('Proyectar bonificaciones adicionales')

    numero_horas_extras_ids = fields.Many2many('hr.payslip.input.type', 'rrhh_num_horas_extras_rel', string='Numero horas extras')
    tipo_entrada_trabajo_id = fields.Many2one('hr.work.entry.type', 'Entrada de trabajo para días trabajados')
    igss_dias_trabajo = fields.Many2one('hr.work.entry.type', 'IGSS días de trabajo')
    centro_trabajo_ids = fields.One2many('rrhh.centro_trabajo', 'company_id', string='Centros de trabajo')
    tipo_planilla_ids = fields.One2many('rrhh.tipo_planilla', 'company_id', string='Tipo de planilla')

    ordinarias_ids = fields.Many2many('hr.salary.rule', 'rrhh_ordinarias_rel', string='Ordinarias')
    extras_ordinarias_ids = fields.Many2many('hr.salary.rule', 'rrhh_extra_ordinarias_rel', string='Extras ordinarias')
    ordinario_ids = fields.Many2many('hr.salary.rule', 'rrhh_ordinario_rel', string='Ordinario')
    extra_ordinario_ids = fields.Many2many('hr.salary.rule', 'rrhh_extra_ordinario_rel', string='Extra ordinario')
    igss_ids = fields.Many2many('hr.salary.rule', 'rrhh_igss_rel', string='IGSS')
    isr_ids = fields.Many2many('hr.salary.rule', 'rrhh_isr_rel', string='ISR')
    anticipos_ids = fields.Many2many('hr.salary.rule', 'rrhh_anticipos_rel', string='Anticipos')
    bonificacion_ids = fields.Many2many('hr.salary.rule', 'rrhh_bonificacion_rel', string='Bonificación incentivo')
    bono_ids = fields.Many2many('hr.salary.rule', 'bono_company_rel', string='Bono 14')
    aguinaldo_ids = fields.Many2many('hr.salary.rule', 'rrhh_aguinaldo_rel', string='Aguinaldo')
    indemnizacion_ids= fields.Many2many('hr.salary.rule', 'rrhh_indemnizacion_rel', string='Retribución por indemnización')
    salario_ids = fields.Many2many('hr.salary.rule', 'rrhh_salario_rel', string='Salario')
    horas_extras_ids = fields.Many2many('hr.salary.rule', 'rrhh_horas_extras_rel', string='Horas extras')
    retribucion_comisiones_ids = fields.Many2many('hr.salary.rule', 'rrhh_retribucion_comisiones_rel', string='Redistribución de comisiones')
    viaticos_ids= fields.Many2many('hr.salary.rule', string='Viaticos')
    retribucion_vacaciones_ids = fields.Many2many('hr.salary.rule', 'rrhh_redistribucion_vacaiones_rel', string='Retribución por vacaciones')
    bonificaciones_adicionales_ids = fields.Many2many('hr.salary.rule', 'rrhh_bonificaciones_adicionales_rel', string='Bonificaciones adicionales')
    extras_ids = fields.Many2many('hr.salary.rule', 'rrhh_extras_rel', string='extras')
    vacaciones_ids = fields.Many2many('hr.salary.rule', 'rrhh_vacaiones_rel', string="Vacaciones")
    decreto_ids = fields.Many2many('hr.salary.rule', 'rrhh_decretro_rel', string="Decreto")
    fija_ids = fields.Many2many('hr.salary.rule', 'rrhh_fija_rel', string="Fija")
    variable_ids = fields.Many2many('hr.salary.rule', 'rrhh_variable_rel', string="Variable")
    otro_salario_ids = fields.Many2many('hr.salary.rule', 'rrhh_otro_salario_rel', string='Otros salarios')
    boni_incentivo_decreto_ids = fields.Many2many('hr.salary.rule', 'rrhh_boni_incentivo_decreto_rel', string='Bonificación incentivo decreto')
    devolucion_isr_otro_ids = fields.Many2many('hr.salary.rule', 'rrhh_dev_isr_otro_rel', string='Devolución ISR')
    salario_total_ids = fields.Many2many('hr.salary.rule', 'rrhh_salario_total_rel', string='Salario total')
    otro_ingreso_afecto_ids = fields.Many2many('hr.salary.rule', 'rrhh_otro_ingresoa_rel', string='Otro ingreso afecto')
    sueldo_igss_ids = fields.Many2many('hr.salary.rule', 'rrhh_sueldo_igss_rel', string='Sueldo IGSS')
    ajuste_ids = fields.Many2many('hr.salary.rule', 'rrhh_ajustes_rel', string='Ajustes ISR')

    # TODO: Quitar todos los siguientes en la siguiente versión
    telefonos = fields.Char('Teléfonos (separados por guiones o diagonales)') # no parece usarse
    fax = fields.Char('Fax') # no parece usarse
    nombre_contacto = fields.Char('Nombre del contacto en centro de trabajo') # no parece usarse
    correo_electronico = fields.Char('correo_electronico') # no parece usarse
    codigo_departamento = fields.Char('Código departamento de la República') # no parece usarse
    codigo_municipio = fields.Char('Código municipio de la República') # no parece usarse
    codigo_actividad_economica = fields.Char('Código actividad económica') # no parece usarse
    identificacion_tipo_planilla = fields.Char('Identificación de tipo de planilla') # no parece usarse
    nombre_tipo_planilla = fields.Char('Nombre del tipo de planilla') # no parece usarse
    tipo_afiliados = fields.Selection([('S', 'Sin IVS'), ('C', 'Con IVS')], 'Tipo de afiliados') # no parece usarse
    periodo_planilla = fields.Selection([('M', 'Mensual'), ('C', 'Catorcenal'), ('S', 'Semanal')], 'Periodo de planilla') # no parece usarse
    departamento_republica = fields.Char('Depto. de la Rep. donde laboran los empleados') # no parece usarse
    actividad_economica = fields.Char('Actividad económica') # no parece usarse
    clase_planilla = fields.Selection([('N', 'Normal'), ('V', 'Sin movimiento')], 'Clase de planilla') # no parece usarse
    codigo_centro_trabajo = fields.Char('Código del centro de trabajo') # no parece usarse
    nombre_centro_trabajo = fields.Char('Nombre del centro de trabajo') # no parece usarse
    direccion_centro_trabajo = fields.Char('Dirección del centro de trabajo') # no parece usarse
    salario_promedio_ids = fields.Many2many('hr.salary.rule','rrhh_salario_promedio_rel', string="Salario promedio") # no parece usarse
    descuentos_ids = fields.Many2many('hr.salary.rule','rrhh_descuentos_rel', string='sescuentos') # no parece usarse
    septimos_asuetos_ids = fields.Many2many('hr.salary.rule','rrhh_septimos_asuetos_rel', string="Séptimos y asuetos") # no parece usarse
    marca = fields.Char('Marca') # no parece usarse
    rango_ingresos = fields.Selection([('si', 'Si'), ('no', 'No')], 'Rango ingresos anual') # no parece usarse
    origen_compania = fields.Selection([('nacional', 'Nacional'), ('extranjero', 'Extranjero')], 'Nacional o Extranjero') # no parece usarse

class rrhh_centro_trabajo(models.Model):
    _name = 'rrhh.centro_trabajo'
    _description = 'Centro de trabajo'
    _rec_name = 'nombre'

    company_id = fields.Many2one('res.company', 'Compañía')
    codigo = fields.Char('Código')
    nombre = fields.Char('Nombre')
    direccion = fields.Char('Dirección')
    zona = fields.Char('Zona')
    telefono = fields.Char('Teléfono')
    fax = fields.Char('Fax')
    nombre_contacto = fields.Char('Nombre contacto')
    correo_electronico = fields.Char('Correo electrónico')
    codigo_departamento = fields.Char('Código departamento')
    codigo_municipio = fields.Char('Código municipio')
    codigo_actividad_economica = fields.Char('Código actividad economica')

class rrhh_tipo_planilla(models.Model):
    _name = 'rrhh.tipo_planilla'
    _description = 'Tipo de planilla'

    company_id = fields.Many2one('res.company', 'Compañía')
    codigo = fields.Char('Código')
    name = fields.Char('Nombre')
    tipo_afiliado = fields.Char('Tipo de afiliados')
    periodo_planilla = fields.Char('Periodo de planilla')
    codigo_departamento = fields.Char('Codigo departamento')
    codigo_actividad_economica = fields.Char('Codigo actividad economica')
    clase_planilla = fields.Char('Clase de planilla')
    tiempo_contrato = fields.Char('Tiempo de contrato')
    liquidaciones_ids = fields.One2many('rrhh.liquidacion_tipo_planilla', 'tipo_planilla_id', string='Liquidación')

class rrhh_liquidacion_tipo_planilla(models.Model):
    _name = 'rrhh.liquidacion_tipo_planilla'
    _description = 'Liquidación del tipo de planilla'

    tipo_planilla_id = fields.Many2one('rrhh.tipo_planilla', 'Tipo de planilla')
    numero = fields.Char('Número')
    complementaria_original = fields.Char('Complementaria u original')
    numero_nota_cargo = fields.Char('Número nota de cargo')
