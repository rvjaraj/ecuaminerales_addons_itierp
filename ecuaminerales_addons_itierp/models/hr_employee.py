# -*- coding: utf-8 -*-
# Part of Odoo. See LICENSE file for full copyright and licensing details.

from odoo import fields, models


class Employee(models.Model):
    _inherit = 'hr.employee'

    codigo_clock = fields.Integer("Código Relog", store=True, help="Código unico asigando en relog")

    _sql_constraints = [
        ('codigo_clock_uniq', 'unique (codigo_clock)', 'El código de relog debe ser unico por empleado'),
    ]
