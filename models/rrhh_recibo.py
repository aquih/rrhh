# -*- coding: utf-8 -*-

from odoo import models, fields, api

class rrhh_recibo(models.Model):
    _name = 'rrhh.recibo'
    _description = 'Recibo'

    name = fields.Char('Nombre', size=40, required=True)
    descripcion = fields.Char('Descripción', size=120)
    linea_id = fields.One2many('rrhh.recibo.regla', 'recibo_id', 'Lineas')
    linea_ingreso_id = fields.One2many('rrhh.recibo.regla', 'recibo_id', 'Ingresos', domain=[('tipo','=','ingreso')], context={'default_tipo':'ingreso'})
    linea_deduccion_id = fields.One2many('rrhh.recibo.regla', 'recibo_id', 'Deducciones', domain=[('tipo','=','deduccion')], context={'default_tipo':'deduccion'})
    linea_entrada_id = fields.One2many('rrhh.recibo.entrada','recibo_id', string='Entradas')

class rrhh_recibo_linea(models.Model):
    _name = 'rrhh.recibo.regla'
    _description = 'Linea de regla'

    name = fields.Char('Nombre', size=40, required=True)
    tipo = fields.Selection([ ('ingreso','Ingreso'), ('deduccion','Deducción') ], 'Tipo')
    sequence = fields.Integer('Secuencia', required=True, index=True, default=5)
    regla_id = fields.Many2many('hr.salary.rule', string='Reglas')
    recibo_id = fields.Many2one('rrhh.recibo', 'Recibo', required=False)

class rrhh_entrada_linea(models.Model):
    _name = 'rrhh.recibo.entrada'
    _description = 'Linea de entrada'

    input_id = fields.Many2one('hr.payslip.input.type',string='Entradas')
    recibo_id = fields.Many2one('rrhh.recibo','Recibo',required=False)
