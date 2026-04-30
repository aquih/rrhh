from odoo import models, fields, api, _
from datetime import date
import logging

class HrEmployeePrivate(models.Model):
    _inherit = 'hr.employee'

    diario_pago_id = fields.Many2one('account.journal', 'Diario de Pago', groups="hr.group_hr_user")
    recibo_id = fields.Many2one('rrhh.recibo', 'Formato de recibo', groups="hr.group_hr_user")
    prestamo_ids = fields.One2many('rrhh.prestamo','employee_id','Prestamo', groups="hr.group_hr_user")
    cantidad_prestamos = fields.Integer(compute='_compute_cantidad_prestamos', string='Prestamos', groups="hr.group_hr_user")
    numero_liquidacion = fields.Char('Numero o identificacion de liquidacion', groups="hr.group_hr_user")
    codigo_centro_trabajo = fields.Char('Codigo de centro de trabajo asignado', groups="hr.group_hr_user")
    codigo_ocupacion = fields.Char('Codigo ocupacion', groups="hr.group_hr_user")
    condicion_laboral = fields.Selection([('P', 'Permanente'), ('T', 'Temporal')], 'Condicion laboral', groups="hr.group_hr_user")
    igss = fields.Char('IGSS', groups="hr.group_hr_user")
    irtra = fields.Char('IRTRA', groups="hr.group_hr_user")
    nivel_academico = fields.Char('Nivel Academico', groups="hr.group_hr_user")
    profesion = fields.Char('Profesion', groups="hr.group_hr_user")
    jornada_trabajo = fields.Char('Jornada de Trabajo', groups="hr.group_hr_user")
    permiso_trabajo = fields.Char('Permiso de Trabajo', groups="hr.group_hr_user")
    edad = fields.Integer(string='Edad', compute="_get_edad", groups="hr.group_hr_user")
    documento_identificacion = fields.Char('Tipo documento identificacion', groups="hr.group_hr_user")
    pueblo_pertenencia = fields.Char('Pueblo de pertenencia', groups="hr.group_hr_user")
    primer_nombre = fields.Char('Primer nombre', groups="hr.group_hr_user")
    segundo_nombre = fields.Char('Segundo nombre', groups="hr.group_hr_user")
    tercer_nombre = fields.Char('Tercer nombre', groups="hr.group_hr_user")
    primer_apellido = fields.Char('Primer apellido', groups="hr.group_hr_user")
    segundo_apellido = fields.Char('Segundo apellido', groups="hr.group_hr_user")
    apellido_casada = fields.Char('Apellido casada', groups="hr.group_hr_user")
    centro_trabajo_id = fields.Many2one('res.company.centro_trabajo',string='Centro de trabajo', groups="hr.group_hr_user")
    tipo_salario = fields.Char('Tipo salario', default="1", groups="hr.group_hr_user")
    tiempo_contrato = fields.Char('Tiempo de contrato', default="TC", groups="hr.group_hr_user")
    nacionalidad = fields.Char('Nacionalidad', groups="hr.group_hr_user")
    tipo_discapacidad = fields.Char('Tipo de discapacidad', groups="hr.group_hr_user")
    comunidad_linguistica = fields.Char('Comunidad Linguística', groups="hr.group_hr_user")
    sucursal = fields.Char("Sucursal",  groups="hr.group_hr_user")
    tipo_contrato = fields.Char("Tipo de contrato",  groups="hr.group_hr_user")

    base_extra = fields.Monetary(readonly=False, related="version_id.base_extra", inherited=True, groups="hr_payroll.group_hr_payroll_user")
    bonificacion_decreto = fields.Monetary(readonly=False, related="version_id.bonificacion_decreto", inherited=True, groups="hr_payroll.group_hr_payroll_user")
    fecha_reinicio_labores = fields.Date(readonly=False, related="version_id.fecha_reinicio_labores", inherited=True, groups="hr_payroll.group_hr_payroll_user")
    temporalidad_contrato = fields.Char(readonly=False, related="version_id.temporalidad_contrato", inherited=True, groups="hr_payroll.group_hr_payroll_user")
    calcula_indemnizacion = fields.Boolean(readonly=False, related="version_id.calcula_indemnizacion", inherited=True, groups="hr_payroll.group_hr_payroll_user")
    historial_salario_ids = fields.One2many(readonly=False, related="version_id.historial_salario_ids", inherited=True, groups="hr_payroll.group_hr_payroll_user")
    
    # TODO: Quitar en la siguiente versión
    codigo_empleado = fields.Char('Código del empleado', groups="hr.group_hr_user")  # usar registration_number
    nit = fields.Char('NIT', groups="hr.group_hr_user") # usar el vat del work_contact_id
    etnia = fields.Char('Etnia', groups="hr.group_hr_user") # no parece usarse
    idioma = fields.Char('Idioma', groups="hr.group_hr_user") # no parece usarse
    pais_origen = fields.Many2one('res.country', 'Pais Origen', groups="hr.group_hr_user") # usar country_of_birth
    codigo_pais_origen = fields.Char('Codigo pais', groups="hr.group_hr_user") # usar country_of_birth
    trabajado_extranjero = fields.Boolean('A trabajado en el extranjero', groups="hr.group_hr_user") # no parece usarse
    motivo_finalizacion = fields.Char('Motivo de finalizacion', groups="hr.group_hr_user") # no parece usarse
    contacto_emergencia = fields.Many2one('res.partner','Contacto de Emergencia', groups="hr.group_hr_user") # usar emergency_contact
    vecindad_dpi = fields.Char('Vecindad DPI', groups="hr.group_hr_user") # no parece usarse
    tarjeta_salud = fields.Boolean('Tarjeta de salud', groups="hr.group_hr_user") # no parece usarse
    tarjeta_manipulacion = fields.Boolean('Tarjeta de manipulación', groups="hr.group_hr_user") # no parece usarse
    tarjeta_pulmones = fields.Boolean('Tarjeta de pulmones', groups="hr.group_hr_user") # no parece usarse
    tarjeta_fecha_vencimiento = fields.Date('Fecha de vencimiento tarjeta de salud', groups="hr.group_hr_user") # no parece usarse
    departamento_id = fields.Many2one('res.country.state','Departmento', groups="hr.group_hr_user") # no parece usarse
    pais_id = fields.Many2one('res.country','Pais', groups="hr.group_hr_user") # no parece usarse
    forma_trabajo_extranjero = fields.Char('Forma trabajada en el extranjero', groups="hr.group_hr_user") # no parece usarse
    pais_trabajo_extranjero_id = fields.Many2one('res.country','Pais trabajado en el extranjero', groups="hr.group_hr_user") # no parece usarse
    finalizacion_laboral_extranjero = fields.Char('Motivo de finalización de la relación laboral en el extranjero', groups="hr.group_hr_user") # no parece usarse

    @api.depends('birthday')
    def _get_edad(self):
        for employee in self:
            if employee.birthday:
                today = date.today()
                employee.edad = today.year - employee.birthday.year - ((today.month, today.day) < (employee.birthday.month, employee.birthday.day))
            else:
                employee.edad = 0

    @api.depends('prestamo_ids')
    def _compute_cantidad_prestamos(self):
        prestamos = self.env['rrhh.prestamo'].search([('employee_id', 'in', self.ids)]).grouped('employee_id')
        for employee in self:
            employee.cantidad_prestamos = len(prestamos.get(employee)) if prestamos.get(employee) else 0
