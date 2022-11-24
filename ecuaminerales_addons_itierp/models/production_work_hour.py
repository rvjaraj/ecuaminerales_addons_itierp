import fileinput
from datetime import datetime, timedelta, date

from odoo import api, fields, models
from odoo.exceptions import ValidationError
import xlrd
import calendar
import xlsxwriter
import io
import base64
from io import BytesIO

# NUMERO MINUTOS DIFERENCIA
MINUTOS_DUPLICADO = 7
TIEMPO_NO_EXTRA = 0.50


class ProductionWorkHour(models.Model):
    _name = 'production.work.hour'
    _description = 'Modulo de Horas en producción '
    _rec_name = 'sequence'
    _order = 'sequence desc'

    sequence = fields.Char('Secuencia', required=False, readonly=True, track_visibility='onchange')
    document = fields.Binary('Documento de Horas', store=True)
    file_name = fields.Char('File Name', track_visibility='onchange')
    state = fields.Selection([('draft', 'Borrador'),
                              ('load', 'Cargado'),
                              ('purify', 'Depurado'),
                              ('confirm', 'Confirmar'),
                              ('posted', 'Pagado')],
                             readonly=True, default='draft', store=True, string="Estado", track_visibility='onchange')
    search_selection = fields.Selection([('code', 'Codigo Reloj'), ('name', 'Nombre')],
                                        default='code', store=True, string="Ubicar empleado por",
                                        track_visibility='onchange')
    hour_production_ids = fields.One2many('production.work.hour.employee', 'production_work_hour', 'Lista de Horas')
    message = fields.Html("Mensaje de Error")
    register_count = fields.Integer('Numero de Registros', compute='_compute_count_registers')
    employee_search = fields.Many2one('hr.employee', 'Empleado')
    fecha_inicio = fields.Datetime('Fecha Inicio', store=True)
    fecha_fin = fields.Datetime('Fecha Fin', store=True)
    number_of_days = fields.Integer('Numero de Dias')
    turnos_rotativos_html = fields.Html('Turnos Rotativos')
    turnos_ocho_horas = fields.Html('8H00-17H00')
    turnos_seguido = fields.Html('6h00-14h00')
    file = fields.Binary('document')
    festivo_start = fields.Datetime('Inicio Feriado')
    festivo_end = fields.Datetime('Fin Feriado')

    hour_festivos_ids = fields.One2many('production.work.hour.festivo', 'parent_id', 'Lista de dias festivos')

    @api.multi
    def delete_dias_festivos(self):
        self.hour_production_ids.write({'festivo': False})
        self.hour_festivos_ids = False

    @api.multi
    def insert_dias_festivos(self):
        if not self.festivo_end or not self.festivo_start:
            raise ValidationError("Ingrese Fechas de Inicio y Fin ")
        for employee_id in set(self.hour_production_ids.mapped('employee_id')):
            list_hours = self.hour_production_ids.filtered(lambda x: x.employee_id == employee_id).sorted('fecha_time')
            count = 1
            for ahora in list_hours[1:]:
                antes = list_hours[count - 1]
                count = count + 1
                if ahora.fecha_time <= self.festivo_end and antes.fecha_time >= self.festivo_start:
                    antes.festivo = True
                    ahora.festivo = True

        festivo_line = self.env['production.work.hour.festivo'].create({
            'parent_id': self.id,
            'festivo_start': self.festivo_start,
            'festivo_end': self.festivo_end})

        self.festivo_start = False
        self.festivo_end = False

    def _compute_count_registers(self):
        self.register_count = len(self.hour_production_ids)

    def view_registro_horas(self):
        self.ensure_one()
        return {
            'name': 'Registro de Horas',
            'type': 'ir.actions.act_window',
            'view_mode': 'tree,form',
            'res_model': 'production.work.hour.employee',
            'context': {'default_production_work_hour': self.id},
            'domain': [('id', 'in', self.hour_production_ids.ids)],
        }

    @api.multi
    def change_to_draft(self):
        self.state = 'draft'

    @api.model
    def create(self, vals):
        vals['sequence'] = self.env['ir.sequence'].next_by_code('production.work.hour.sequence') or '/'
        return super(ProductionWorkHour, self).create(vals)

    def load_information_of_file(self):
        if not self.document:
            raise ValidationError("Cargar Archivo de Horas")
        wb = xlrd.open_workbook(file_contents=base64.decodestring(self.document))
        sheet = wb.sheets()[0] if wb.sheets() else None
        data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
        if data[0] == ['', '', '', '', '', '']:
            data.remove(data[0])
        # if data[0] != ['Nombre', 'Número de empleado', 'Departamento', 'Fecha', 'Hora', 'Dispositivo']:
        #     raise ValidationError("""Recurde que el archivo debe contener la siguiente estructura  \n
        #     ['Nombre', 'Número de empleado', 'Departamento', 'Fecha', 'Hora', 'Dispositivo']""")
        self.hour_production_ids = False
        names_no_search = []
        name_not_range = []
        data_create = []
        for count, line in enumerate(data[1:]):
            name = line[0]
            employee = []
            if self.search_selection == 'name':
                employee = self.env['hr.employee'].search([('name', '=', line[0])])
            if self.search_selection == 'code':
                employee = self.env['hr.employee'].search([('codigo_clock', '=', int(line[1]))])
            if not employee:
                names_no_search.append(name)
                continue
            if employee and len(employee) > 1:
                employee = employee[0]
                name_not_range.append(name)
            data_create.append((0, count, {'employee_id': employee.id,
                                           'departamento': line[2],
                                           'fecha_time': self.conv_date_hout(line[3], line[4]),
                                           'hour': self.conv_time_float(line[4]),
                                           'dispositivo': line[5]}))
        self.hour_production_ids = data_create
        self.insert_messages(name_not_range, names_no_search)
        self.purge_data()
        self.state = 'load'

    def insert_messages(self, name_not_range, names_no_search):
        self.message = False
        if name_not_range or names_no_search:
            self.message = "<ul>"
        else:
            self.message = False
        for name in set(name_not_range):
            self.message += "<li class='text-danger'> Empleado: %s multiples coincidencias </li> \n " % name
        for name in set(names_no_search):
            self.message += "<li class='text-warning'>  Empleado: %s no encontrado </li> \n " % name

        if self.message:
            self.message += "</ul>"

    def conv_time_float(self, value):
        vals = value.split(':')
        t, hours = divmod(float(vals[0]), 24)
        t, minutes = divmod(float(vals[1]), 60)
        minutes = minutes / 60.0
        return float(hours + minutes)

    def conv_date_hout(self, date, time):
        date_time_str = date + " " + time + ':00'
        fecha = datetime.strptime(str(date_time_str), '%m/%d/%Y %H:%M:%S')
        fecha = fecha + timedelta(hours=5)
        return fecha

    def purge_data(self):
        if not self.hour_production_ids:
            return True
        self.hour_production_ids.write({'delete': False, 'type_mar': 'error', 'dif': 0, 'turno': 'no'})
        for employee_id in set(self.hour_production_ids.mapped('employee_id')):
            list_hours = self.hour_production_ids.filtered(lambda x: x.employee_id == employee_id).sorted('fecha_time')
            count = 1
            for ahora in list_hours[1:]:
                antes = list_hours[count - 1]
                diferencia = ahora.fecha_time - antes.fecha_time
                minutes = abs(diferencia.total_seconds() / 60)
                ahora.dif = minutes
                ahora.dif_h = minutes / 60
                if minutes < MINUTOS_DUPLICADO:
                    ahora.delete = True
                    ahora.type_mar = antes.type_mar
                    ahora.turno = antes.turno
                    if ahora.type_mar == 'exit':
                        antes.delete = True
                        ahora.delete = False
                else:
                    self.detectar_ingreso_salida(antes, ahora, minutes)
                count += 1
        if self.hour_production_ids:
            self.fecha_inicio = min(self.hour_production_ids.mapped('fecha_time'))
            self.fecha_fin = max(self.hour_production_ids.mapped('fecha_time'))
            self.number_of_days = (self.fecha_fin - self.fecha_inicio).days
        turno = self.env.ref('ecuaminerales_addons_itierp.resource_rotativos')
        marcaciones_old = self.hour_production_ids.filtered(
            lambda x: x.type_mar in ['old', 'error'] and x.resource_calendar_id == turno)
        marcaciones1 = marcaciones_old.filtered(
            lambda x: (x.fecha_time - timedelta(hours=5)).strftime('%d-%m-%y') == self.fecha_inicio.strftime(
                '%d-%m-%y'))
        marcaciones1.write({'delete': True})
        marcaciones2 = marcaciones_old.filtered(
            lambda x: x.fecha_time - timedelta(hours=5) >= self.fecha_fin - timedelta(days=1, hours=5) and x.dif_h >= 3)
        if not self.hour_production_ids.filtered(lambda x: x.dispositivo == 'CREADO') and marcaciones2:
            for employee_id in marcaciones2.mapped('employee_id'):
                for item in marcaciones2.filtered(lambda x: x.employee_id == employee_id).sorted('fecha_time')[-1]:
                    item_cp = item.copy({'dispositivo': "CREADO"})
                    item_cp.fecha_time = item_cp.fecha_time + timedelta(hours=8, minutes=10)
            self.purge_data()
        creado = self.hour_production_ids.filtered(
            lambda x: x.type_mar in ['error'] and x.dispositivo == 'CREADO')
        if creado:
            creado.unlink()

    def detectar_ingreso_salida(self, antes, ahora, minutes):
        f_antes = antes.fecha_time - timedelta(hours=5)
        f_ahora = ahora.fecha_time - timedelta(hours=5)
        if antes.turno != 'no':
            return True
        # TURNO ROTATIVOS
        sales_journal_id = self.env.ref('ecuaminerales_addons_itierp.resource_rotativos')
        if ahora.resource_calendar_id == sales_journal_id:
            # VALIDAR HORARIOS
            # 	             |      TURNO 1	   | TURNO 2	        | TURNO 3
            # LUNES-VIERNES	 | 06H00 - 14H00   |	14H00 - 22H00	| 22H00 - 06H00
            # SÁBADO	     |06H00 - 18H00	   |    LIBRE	        | 18H00 - 06H00
            # DOMINGO	     |06H00 - 18H00    |	18H00 - 06H00	| LIBRE
            if f_antes.weekday() in [calendar.MONDAY, calendar.TUESDAY, calendar.WEDNESDAY, calendar.THURSDAY,
                                     calendar.FRIDAY]:
                if (minutes / 60) > 14:
                    antes.type_mar = 'old'
                    return True
                if 5 <= f_antes.hour <= 8 and f_ahora.hour <= 18:
                    antes.type_mar = 'income'
                    ahora.type_mar = 'exit'
                    antes.turno = 't1'
                    ahora.turno = 't1'
                    return True
                if 13 <= f_antes.hour <= 15 and f_ahora.hour <= 24 or 13 <= f_antes.hour <= 15 and f_ahora.hour <= 2:
                    antes.type_mar = 'income'
                    ahora.type_mar = 'exit'
                    antes.turno = 't2'
                    ahora.turno = 't2'
                    return True
                if 21 <= f_antes.hour <= 23 and f_ahora.hour <= 8:
                    antes.type_mar = 'income'
                    ahora.type_mar = 'exit'
                    antes.turno = 't3'
                    ahora.turno = 't3'
                    return True
                if 9 <= f_antes.hour <= 11 and f_ahora.hour <= 23:
                    antes.type_mar = 'income'
                    ahora.type_mar = 'exit'
                    antes.turno = 'tt2'
                    ahora.turno = 'tt2'
                    return True

            if f_antes.weekday() in [calendar.SATURDAY, calendar.SUNDAY]:
                if 5 <= f_antes.hour <= 7 and f_ahora.hour <= 20 and f_ahora.day == f_antes.day:
                    antes.type_mar = 'income'
                    ahora.type_mar = 'exit'
                    antes.turno = 't1f'
                    ahora.turno = 't1f'
                    return True

            if f_antes.weekday() in [calendar.SATURDAY]:
                if 15 <= f_antes.hour <= 19 and f_ahora.hour <= 8:
                    antes.type_mar = 'income'
                    ahora.type_mar = 'exit'
                    antes.turno = 't2f'
                    ahora.turno = 't2f'
                    return True
            if f_antes.weekday() in [calendar.SUNDAY]:
                if 15 <= f_antes.hour <= 19 and f_ahora.hour <= 8:
                    antes.type_mar = 'income'
                    ahora.type_mar = 'exit'
                    antes.turno = 't3f'
                    ahora.turno = 't3f'
                    return True
            return False
        calendar_almuerzo = self.env.ref('ecuaminerales_addons_itierp.resource_ocho_horas_1_almuerzo')
        if ahora.resource_calendar_id == calendar_almuerzo:
            if 5 <= f_antes.hour <= 8 and f_ahora.hour <= 14:
                antes.type_mar = 'income'
                ahora.type_mar = 'exit'
                antes.turno = 'morning'
                ahora.turno = 'morning'
                return True
            if 12 <= f_antes.hour <= 15 and f_ahora.hour >= 15:
                antes.type_mar = 'income'
                ahora.type_mar = 'exit'
                antes.turno = 'late'
                ahora.turno = 'late'
                return True
            if 5 <= f_antes.hour <= 8 and f_ahora.hour >= 16:
                antes.type_mar = 'income'
                ahora.type_mar = 'exit'
                antes.turno = 'morning'
                ahora.turno = 'late'
                return True
            if (minutes / 60) > 9:
                if 15 <= antes.hour >= 16:
                    antes.turno = 'late'
                if 5 <= antes.hour >= 9:
                    antes.turno = 'morning'
                antes.type_mar = 'old'
                return True
        resource_5h_14_h = self.env.ref('ecuaminerales_addons_itierp.resource_5h_14_h')
        if ahora.resource_calendar_id == resource_5h_14_h:
            if 4 <= f_antes.hour <= 7 and f_ahora.hour >= 14:
                antes.type_mar = 'income'
                ahora.type_mar = 'exit'
                antes.turno = 'seguido'
                ahora.turno = 'seguido'
                return True

            if (minutes / 60) > 10:
                antes.turno = 'seguido'
                antes.type_mar = 'old'
                return True

    def turnos_rotativos_html_insertion(self):
        if not self.hour_production_ids:
            self.turnos_rotativos_html = ""
            return True
        sales_journal_id = self.env.ref('ecuaminerales_addons_itierp.resource_rotativos')
        data_filter = self.hour_production_ids.filtered(
            lambda x: x.resource_calendar_id == sales_journal_id and x.turno != 'no')
        if not data_filter:
            self.turnos_rotativos_html = ""
            return True

        html_text = """<table class="o_list_view table table-sm table-hover table-striped o_list_view_ungrouped">
                                        <thead>
                                        <tr><th>Empleado</th>
                                        """
        fecha_header = self.fecha_inicio.strftime('%d-%m')
        for day in range(self.number_of_days):
            html_text += """<th>%s</th>""" % fecha_header
            fecha_header = (self.fecha_inicio + timedelta(days=day + 1)).strftime('%d-%m')

        html_text += """</thead>"""
        for employee_id in data_filter.mapped('employee_id').sorted('name'):
            list_hours = data_filter.filtered(lambda x: x.employee_id == employee_id).sorted('fecha_time')
            html_text += """<tr><th>%s</th>""" % employee_id.display_name
            fecha_trabajo = self.fecha_inicio.strftime('%d-%m-%y')
            for day in range(1, self.number_of_days + 1):
                data = list_hours.filtered(lambda x: x.fecha_time.strftime('%d-%m-%y') == fecha_trabajo)
                if data:
                    inicio = max(data.mapped('fecha_time')) - min(data.mapped('fecha_time'))
                    html_text += """<th>%s</th>""" % round(inicio.total_seconds() / 60 / 60, 2)
                else:
                    html_text += """<th class="text-danger">X</th>"""
                fecha_trabajo = (self.fecha_inicio + timedelta(days=day)).strftime('%d-%m-%y')
            html_text += """</tr>"""
        self.turnos_rotativos_html = html_text + """</tbody></table>"""

    def _get_days_header(self, data_filter):
        days = data_filter.sorted('fecha_time').mapped('fecha_time')
        days = [(x - timedelta(hours=5)).replace(hour=0, second=0, minute=0) for x in days]
        return sorted(set(days))

    def _get_data_filter(self):
        ocho_horas_cal = self.env.ref('ecuaminerales_addons_itierp.resource_ocho_horas_1_almuerzo')
        return self.hour_production_ids.filtered(
            lambda x: x.resource_calendar_id == ocho_horas_cal and x.turno != 'no')

    def _get_data_filter_seguido(self):
        ocho_horas_cal = self.env.ref('ecuaminerales_addons_itierp.resource_5h_14_h')
        return self.hour_production_ids.filtered(
            lambda x: x.resource_calendar_id == ocho_horas_cal and x.turno != 'no')

    def turnos_ocho_horas_html_insertion(self):
        if not self.hour_production_ids:
            self.turnos_ocho_horas = ""
            return True
        data_filter = self._get_data_filter()
        if not data_filter:
            self.turnos_ocho_horas = ""
            return True

        html_text = """<table class="o_list_view table table-sm table-hover table-striped o_list_view_ungrouped">
                                        <thead>
                                        <tr><th>Empleado</th>
                                        """
        days = self._get_days_header(data_filter)
        for day in days:
            html_text += """<th>%s</th>""" % day.strftime('%d-%m')
        html_text += """</thead>"""
        for employee_id in data_filter.mapped('employee_id').sorted('name'):
            html_text += """<tr><th>%s</th>""" % employee_id.display_name
            list_hours = data_filter.filtered(lambda x: x.employee_id == employee_id).sorted('fecha_time')
            for day in days:
                horas = list_hours.filtered(
                    lambda x: (x.fecha_time - timedelta(hours=5)).strftime('%d-%m') == day.strftime('%d-%m'))
                horas = horas.sorted('fecha_time')
                if len(horas) == 4:
                    h1 = horas[1].fecha_time - horas[0].fecha_time
                    h1 += horas[3].fecha_time - horas[2].fecha_time
                    html_text += """<th class="text-success">%s</th>""" % round(h1.total_seconds() / 60 / 60, 2)
                elif len(horas) == 2:
                    h1 = (horas[1].fecha_time - timedelta(hours=1)) - horas[0].fecha_time
                    html_text += """<th class="text-info">%s</th>""" % round(h1.total_seconds() / 60 / 60, 2)
                elif horas:
                    h1 = horas[-1].fecha_time - horas[0].fecha_time
                    html_text += """<th class="text-warning">%s</th>""" % round(h1.total_seconds() / 60 / 60, 2)
                else:
                    html_text += """<th class="text-danger">X</th>"""
        self.turnos_ocho_horas = html_text + """</tbody></table>"""

    def turnos_seguido_html_insertion(self):
        if not self.hour_production_ids:
            self.turnos_seguido = ""
            return True
        data_filter = self._get_data_filter_seguido()
        if not data_filter:
            self.turnos_seguido = ""
            return True
        html_text = """<table class="o_list_view table table-sm table-hover table-striped o_list_view_ungrouped">
                                        <thead>
                                        <tr><th>Empleado</th>
                                        """
        days = self._get_days_header(data_filter)
        for day in days:
            html_text += """<th>%s</th>""" % day.strftime('%d-%m')
        html_text += """</thead>"""
        for employee_id in data_filter.mapped('employee_id').sorted('name'):
            html_text += """<tr><th>%s</th>""" % employee_id.display_name
            list_hours = data_filter.filtered(lambda x: x.employee_id == employee_id).sorted('fecha_time')
            for day in days:
                horas = list_hours.filtered(
                    lambda x: (x.fecha_time - timedelta(hours=5)).strftime('%d-%m') == day.strftime('%d-%m'))
                horas = horas.sorted('fecha_time')
                if len(horas) == 2:
                    h1 = horas[1].fecha_time - horas[0].fecha_time
                    html_text += """<th>%s</th>""" % round(h1.total_seconds() / 60 / 60, 2)
                elif horas:
                    h1 = horas[-1].fecha_time - horas[0].fecha_time
                    html_text += """<th class="text-warning">%s</th>""" % round(h1.total_seconds() / 60 / 60, 2)
                else:
                    html_text += """<th class="text-danger">X</th>"""
        self.turnos_seguido = html_text + """</tbody></table>"""

    def delete_duplicates(self):
        self.hour_production_ids = self.hour_production_ids.filtered(lambda x: not x.delete)
        self.purge_data()
        self.state = 'purify'
        self.turnos_rotativos_html_insertion()
        self.turnos_ocho_horas_html_insertion()
        self.turnos_seguido_html_insertion()

    def print_header_excel(self, sheet, format_center):
        sheet.set_column(0, 0, 45)
        sheet.merge_range(0, 0, 1, 0, "Empleado", format_center)
        fecha_header = self.fecha_inicio.strftime('%d-%m')
        count = 1
        for day in range(self.number_of_days + 1):
            sheet.merge_range(0, count, 0, count + 3, fecha_header, format_center)
            sheet.set_column(count, count + 3, 3)
            sheet.write(1, count, "T1", format_center)
            sheet.write(1, count + 1, "T2", format_center)
            sheet.write(1, count + 2, "T3", format_center)
            sheet.write(1, count + 3, "TD", format_center)
            fecha_header = (self.fecha_inicio + timedelta(days=day + 1)).strftime('%d-%m')
            count += 4
        sheet.merge_range(0, count, 1, count, "TOTAL", format_center)
        count += 1
        sheet.merge_range(0, count, 1, count, "T1", format_center)
        count += 1
        sheet.merge_range(0, count, 1, count, "T2", format_center)
        count += 1
        sheet.merge_range(0, count, 1, count, "T3", format_center)
        count += 1
        sheet.merge_range(0, count, 1, count, "TD", format_center)
        count += 1
        sheet.merge_range(0, count, 1, count, "EXTRAS", format_center)

    def excel_turnos_rotativos(self, sheet, format_center):
        sales_journal_id = self.env.ref('ecuaminerales_addons_itierp.resource_rotativos')
        data_filter = self.hour_production_ids.filtered(
            lambda x: x.resource_calendar_id == sales_journal_id and x.turno != 'no')
        fila = 2
        for employee_id in data_filter.mapped('employee_id').sorted('name'):

            sheet.write(fila, 0, employee_id.display_name, format_center)
            list_hours = data_filter.filtered(lambda x: x.employee_id == employee_id).sorted('fecha_time')
            fecha_header = self.fecha_inicio - timedelta(hours=7)
            col = 1
            t1 = 0
            t2 = 0
            t3 = 0
            td = 0
            extras = 0
            for day in range(1, self.number_of_days + 2):
                fecha_nex = fecha_header + timedelta(hours=16)
                data = self.filter_data_turno(list_hours, fecha_header, fecha_nex, ['t1'])
                if data:
                    h, ex = self.print_data_lina_t1_t2(data, col, fila, sheet, 't1')
                    t1 += h
                    col += 1
                else:
                    data = self.filter_data_turno(list_hours, fecha_header, fecha_nex, ['t1f'])
                    h, ex = self.print_data_lina_t1_t2(data, col, fila, sheet, 't1f')
                    t1 += h
                    col += 1

                fecha_header = fecha_header.replace(hour=11, minute=30, second=0)
                fecha_nex = fecha_header + timedelta(hours=16)
                data = self.filter_data_turno(list_hours, fecha_header, fecha_nex, ['t2'])
                if data:
                    h, ex = self.print_data_lina_t1_t2(data, col, fila, sheet, 't2')
                    extras += ex
                    t2 += h
                    col += 1
                else:
                    fecha_header = fecha_header.replace(hour=15, minute=30, second=0)
                    fecha_nex = fecha_header + timedelta(hours=16)
                    data = self.filter_data_turno(list_hours, fecha_header, fecha_nex, ['t2f'])
                    h, ex = self.print_data_lina_t1_t2(data, col, fila, sheet, 't2f')
                    extras += ex
                    t2 += h
                    col += 1

                fecha_header = fecha_header.replace(hour=19, minute=30, second=0)
                fecha_nex = fecha_header + timedelta(hours=16)
                data = self.filter_data_turno(list_hours, fecha_header, fecha_nex, ['t3'])
                if data:
                    h, ex = self.print_data_lina_t1_t2(data, col, fila, sheet, 't3')
                    extras += ex
                    t3 += h
                    col += 1
                else:
                    fecha_header = fecha_header.replace(hour=15, minute=30, second=0)
                    fecha_nex = fecha_header + timedelta(hours=16)
                    data = self.filter_data_turno(list_hours, fecha_header, fecha_nex, ['t3f'])
                    h, ex = self.print_data_lina_t1_t2(data, col, fila, sheet, 't3f')
                    extras += ex
                    t3 += h
                    col += 1

                fecha_header = fecha_header.replace(hour=6, minute=30, second=0)
                fecha_nex = fecha_header + timedelta(hours=19)
                data = self.filter_data_turno(list_hours, fecha_header, fecha_nex, ['tt2'])
                h, ex = self.print_data_lina_t1_t2(data, col, fila, sheet, 'tt2')
                extras += ex
                td += h
                col += 1
                fecha_header = self.fecha_inicio + timedelta(days=day)
                fecha_header = fecha_header - timedelta(hours=7)
            sheet.write(fila, col, t1 + t2 + t3 + td)
            col += 1
            sheet.write(fila, col, t1)
            col += 1
            sheet.write(fila, col, t2)
            col += 1
            sheet.write(fila, col, t3)
            col += 1
            sheet.write(fila, col, td)
            col += 1
            sheet.write(fila, col, extras)
            col += 1
            fila += 1

    def filter_data_turno(self, list_hours, fecha_header, fecha_nex, turnos):
        return list_hours.filtered(
            lambda x: fecha_header <= x.fecha_time - timedelta(hours=5) <= fecha_nex and x.turno in turnos).sorted(
            'fecha_time')

    def get_horas_extras(self, turno, horas):
        if turno in ['t1', 't2', 't3']:
            extra = horas - 8
            if extra > 0 and extra >= TIEMPO_NO_EXTRA:
                return extra
            return 0
        if turno in ['t1f', 't2f', 't3f', 'tt2']:
            extra = horas - 12.50
            if turno in ['tt2']:
                extra = horas - 12
            if extra > 0 and extra >= TIEMPO_NO_EXTRA:
                return extra
            return 0
        return 0

    def print_data_lina_t1_t2(self, data, col, fila, sheet, turno):
        if data:
            inicio = False
            fin = False
            for mark in data.sorted('fecha_time'):
                if mark.type_mar == 'income' and not inicio:
                    inicio = mark
                if mark.type_mar == 'exit' and not fin and inicio:
                    fin = mark
                if inicio and fin:
                    break
            if inicio and fin and inicio.fecha_time < fin.fecha_time:
                horas = fin.fecha_time - inicio.fecha_time
                horas = round(horas.total_seconds() / 60 / 60, 2)
                sheet.write(fila, col, horas)
                return horas, self.get_horas_extras(turno, horas)
            else:
                sheet.write(fila, col, '')
            return 0, 0
        else:
            sheet.write(fila, col, '')
            return 0, 0

    def get_horas_extras_hora(self, horas):
        horas = round(horas.total_seconds() / 60 / 60, 2)
        extra = horas - 8
        if extra < TIEMPO_NO_EXTRA:
            extra = 0
        return horas, extra

    def print_header_almuerzo_excel(self, sheet, format_center):
        sheet.set_column(0, 0, 45)
        sheet.write(0, 0, "Empleado", format_center)
        data_filter = self._get_data_filter()
        days = self._get_days_header(data_filter)
        count = 1
        for day in days:
            sheet.set_column(count, count, 5)
            sheet.write(0, count, day.strftime('%d-%m'), format_center)
            count += 1
        sheet.write(0, count, "TOTAL", format_center)
        count += 1
        sheet.write(0, count, "EXTRAS", format_center)

    def print_header_seguido_excel(self, sheet, format_center):
        sheet.set_column(0, 0, 45)
        sheet.write(0, 0, "Empleado", format_center)
        data_filter = self._get_data_filter_seguido()
        days = self._get_days_header(data_filter)
        count = 1
        for day in days:
            sheet.write(0, count, day.strftime('%d-%m'), format_center)
            sheet.set_column(count, count + 1, 5)
            count += 1
        sheet.write(0, count, "TOTAL", format_center)
        count += 1
        sheet.write(0, count, "EXTRAS", format_center)

    def excel_turnos_almuerzo(self, sheet, format_center):
        data_filter = self._get_data_filter()
        days = self._get_days_header(data_filter)
        fila = 1
        for employee_id in data_filter.mapped('employee_id').sorted('name'):
            col = 0
            sheet.write(fila, col, employee_id.display_name, format_center)
            list_hours = data_filter.filtered(lambda x: x.employee_id == employee_id).sorted('fecha_time')
            total = 0
            ex = 0
            for day in days:
                col += 1
                horas = list_hours.filtered(
                    lambda x: (x.fecha_time - timedelta(hours=5)).strftime('%d-%m') == day.strftime('%d-%m'))
                if len(horas) == 4:
                    h1 = horas[1].fecha_time - horas[0].fecha_time
                    h1 += horas[3].fecha_time - horas[2].fecha_time
                    horas, extra = self.get_horas_extras_hora(h1)
                    total += horas
                    ex += extra
                    sheet.write(fila, col, horas)
                elif len(horas) == 2:
                    h1 = (horas[1].fecha_time - timedelta(hours=5)) - horas[0].fecha_time
                    horas, extra = self.get_horas_extras_hora(h1)
                    total += horas
                    ex += extra
                    sheet.write(fila, col, horas)
                elif horas:
                    h1 = horas[-1].fecha_time - horas[0].fecha_time
                    horas, extra = self.get_horas_extras_hora(h1)
                    total += horas
                    ex += extra
                    sheet.write(fila, col, horas)
                else:
                    sheet.write(fila, col, "")
            col += 1
            sheet.write(fila, col, total)
            col += 1
            sheet.write(fila, col, ex)
            fila += 1

    def excel_turnos_seguido(self, sheet, format_center):
        data_filter = self._get_data_filter_seguido()
        days = self._get_days_header(data_filter)
        fila = 1
        for employee_id in data_filter.mapped('employee_id').sorted('name'):
            col = 0
            sheet.write(fila, col, employee_id.display_name, format_center)
            list_hours = data_filter.filtered(lambda x: x.employee_id == employee_id).sorted('fecha_time')
            total = 0
            ex = 0
            for day in days:
                col += 1
                horas = list_hours.filtered(
                    lambda x: (x.fecha_time - timedelta(hours=5)).strftime('%d-%m') == day.strftime('%d-%m'))
                if len(horas) == 2:
                    h1 = horas[1].fecha_time - horas[0].fecha_time
                    horas, extra = self.get_horas_extras_hora(h1)
                    total += horas
                    ex += extra
                    sheet.write(fila, col, horas)
                elif horas:
                    h1 = horas[-1].fecha_time - horas[0].fecha_time
                    horas, extra = self.get_horas_extras_hora(h1)
                    total += horas
                    ex += extra
                    sheet.write(fila, col, horas)
                else:
                    sheet.write(fila, col, "")
            col += 1
            sheet.write(fila, col, total)
            col += 1
            sheet.write(fila, col, ex)
            fila += 1

    @api.multi
    def print_excel_report(self):
        fp = BytesIO()
        workbook = xlsxwriter.Workbook(fp)
        sheet = workbook.add_worksheet('Turnos Rotativos')
        format_center = workbook.add_format({'bold': True, 'align': 'vcenter'})
        self.print_header_excel(sheet, format_center)
        self.excel_turnos_rotativos(sheet, format_center)
        sheet1 = workbook.add_worksheet('8H00-17H00')
        self.print_header_almuerzo_excel(sheet1, format_center)
        self.excel_turnos_almuerzo(sheet1, format_center)
        sheet2 = workbook.add_worksheet('6h00-14h00')
        self.print_header_seguido_excel(sheet2, format_center)
        self.excel_turnos_seguido(sheet2, format_center)
        return self.return_exel_report(fp, workbook)

    @api.multi
    def print_excel_report_resumen_del(self):
        fp = BytesIO()
        workbook = xlsxwriter.Workbook(fp)
        sheet = workbook.add_worksheet('Turno Rotativos')
        format_center = workbook.add_format({'bold': True, 'align': 'vcenter'})
        col = 0
        sheet.set_column(col, col, 50)
        sheet.write(0, col, "EMPLEADO", format_center)
        col += 1
        sheet.set_column(col, col, 14)
        sheet.write(0, col, "TURNO", format_center)
        col += 1
        sheet.set_column(col, col, 14)
        sheet.write(0, col, "DIA", format_center)
        col += 1
        sheet.set_column(col, col, 14)
        sheet.write(0, col, "INGRESO", format_center)
        col += 1
        sheet.set_column(col, col, 14)
        sheet.write(0, col, "SALIDA", format_center)
        col += 1
        sheet.set_column(col, col, 14)
        sheet.write(0, col, "TRABAJO", format_center)
        col += 1
        sheet.set_column(col, col, 14)
        sheet.write(0, col, "HORAS", format_center)
        col += 1
        sheet.set_column(col, col, 14)
        sheet.write(0, col, "NOCTURNAS", format_center)
        col += 1
        sheet.set_column(col, col, 16)
        sheet.write(0, col, "SUPLEMENTARIAS", format_center)
        col += 1
        sheet.set_column(col, col, 17)
        sheet.write(0, col, "EXTRAORDINARIAS", format_center)
        col += 1
        if not self.hour_production_ids:
            return True
        list_data = self.hour_production_ids.filtered(lambda x: x.type_mar in ['exit', 'income'])
        jornada = self.env.ref('ecuaminerales_addons_itierp.resource_rotativos')
        list_data = list_data.filtered(lambda x: x.resource_calendar_id == jornada)

        fila = 1
        for employee_id in list_data.mapped('employee_id').sorted('name'):
            list_hours = list_data.filtered(lambda x: x.employee_id == employee_id).sorted('fecha_time')
            count = 1
            for ahora in list_hours[1:].sorted('fecha_time'):
                if ahora.type_mar == 'income':
                    count += 1
                    continue
                col = 0
                sheet.write(fila, col, employee_id.display_name)
                antes = list_hours[count - 1]
                f_antes = antes.fecha_time - timedelta(hours=5)
                f_ahora = ahora.fecha_time - timedelta(hours=5)
                diferencia = f_ahora - f_antes
                horas = round(diferencia.total_seconds() / 60 / 60, 2)
                col += 1
                sheet.write(fila, col, ahora.turno)
                col += 1
                sheet.write(fila, col, f_antes.strftime('%d-%m-%Y'))
                col += 1
                sheet.write(fila, col, f_antes.strftime('%H:%M:%S'))
                col += 1
                sheet.write(fila, col, f_ahora.strftime('%H:%M:%S'))
                col += 1
                sheet.write(fila, col, horas)
                trabajo = horas
                extra = 0
                if ahora.turno in ['t1', 't2', 't3']:
                    tiempo_no_ocho = horas - 8
                    if 0 <= tiempo_no_ocho >= TIEMPO_NO_EXTRA:
                        trabajo = horas
                    else:
                        trabajo = 8
                if ahora.turno in ['tt2', 't1f', 't2f', 't3f']:
                    tiempo_no_ocho = horas - 12.50
                    if ahora.turno in ['tt2']:
                        tiempo_no_ocho = horas - 12
                    if 0 <= tiempo_no_ocho >= TIEMPO_NO_EXTRA:
                        trabajo = horas
                    else:
                        trabajo = 11.5
                        if ahora.turno in ['tt2']:
                            trabajo = 12
                trabajo = round(trabajo, 2)
                col += 1
                sheet.write(fila, col, trabajo)
                nocturna = 0
                if ahora.turno == 't2' and not ahora.festivo:
                    f_aux = f_antes.replace(hour=19, minute=30, second=0)
                    horas_n = (f_ahora - f_aux).total_seconds() / 60 / 60
                    if 0 < horas_n - 2.5 >= TIEMPO_NO_EXTRA:
                        nocturna += round(horas_n, 2)
                    else:
                        nocturna += 2.5
                if ahora.turno == 't3' and not ahora.festivo:
                    f_aux = f_ahora.replace(hour=5, minute=30, second=0)
                    horas_n = (f_aux - f_antes).total_seconds() / 60 / 60
                    if 0 < horas_n - 7.5 >= TIEMPO_NO_EXTRA:
                        nocturna += round(horas_n, 2)
                    else:
                        nocturna += 7.5
                col += 1
                sheet.write(fila, col, round(nocturna, 2))
                suplementaria = 0
                if ahora.turno in ['tt2'] and not ahora.festivo:
                    suple = horas - 8
                    if 0 < suple - 4 >= TIEMPO_NO_EXTRA:
                        suplementaria = round(suple, 2)
                    else:
                        suplementaria = 4
                if ahora.turno in ['t1'] and not ahora.festivo:
                    suple = horas - 8
                    if 0 < suple >= TIEMPO_NO_EXTRA:
                        suplementaria = round(suple, 2)
                    else:
                        suplementaria = 0
                col += 1
                sheet.write(fila, col, round(suplementaria, 2))
                extraordinaria = 0
                if ahora.turno in ['t1f', 't2f', 't3f'] or ahora.festivo or antes.festivo:
                    if ahora.festivo or antes.festivo:
                        extraordinaria += trabajo
                    else:
                        ex = horas - 12.50
                        if ex <= 0:
                            extraordinaria += 11.50
                        elif ex >= TIEMPO_NO_EXTRA:
                            extraordinaria += round(horas - 0.50, 2)
                        else:
                            extraordinaria += 12
                            if ahora.turno in ['tt2']:
                                extraordinaria = 12
                col += 1
                sheet.write(fila, col, round(extraordinaria, 2))
                fila += 1
                count += 1

        sheet = workbook.add_worksheet('ADMINISTRACION')
        format_center = workbook.add_format({'bold': True, 'align': 'vcenter'})
        col = 0
        sheet.set_column(col, col, 50)
        sheet.write(0, col, "EMPLEADO", format_center)
        col += 1
        sheet.set_column(col, col, 14)
        sheet.write(0, col, "DIA", format_center)
        col += 1
        sheet.set_column(col, col, 14)
        sheet.write(0, col, "INGRESO", format_center)
        col += 1
        sheet.set_column(col, col, 14)
        sheet.write(0, col, "SALIDA", format_center)
        col += 1
        sheet.set_column(col, col, 14)
        sheet.write(0, col, "TRABAJO", format_center)
        col += 1
        sheet.set_column(col, col, 14)
        sheet.write(0, col, "HORAS", format_center)
        col += 1
        sheet.set_column(col, col, 14)
        sheet.write(0, col, "NOCTURNAS", format_center)
        col += 1
        sheet.set_column(col, col, 16)
        sheet.write(0, col, "SUPLEMENTARIAS", format_center)
        col += 1
        sheet.set_column(col, col, 17)
        sheet.write(0, col, "EXTRAORDINARIAS", format_center)
        col += 1
        data_filter = self._get_data_filter()
        data_filter |= self._get_data_filter_seguido()
        data_filter = data_filter.filtered(lambda x: x.type_mar in ['exit', 'income'])
        days = self._get_days_header(data_filter)
        fila = 1
        for employee_id in data_filter.mapped('employee_id').sorted('name'):
            list_hours = data_filter.filtered(lambda x: x.employee_id == employee_id).sorted('fecha_time')
            for day in days:
                horas = list_hours.filtered(
                    lambda x: (x.fecha_time - timedelta(hours=5)).strftime('%d-%m') == day.strftime('%d-%m'))
                if not horas or len(horas) == 1:
                    continue
                horas.sorted('fecha_time')
                f_antes = horas[0].fecha_time - timedelta(hours=5)
                f_ahora = horas[-1].fecha_time - timedelta(hours=5)
                diferencia = f_ahora - f_antes
                horas_ = round(diferencia.total_seconds() / 60 / 60, 2)
                col = 0
                sheet.write(fila, col, employee_id.display_name)
                col += 1
                sheet.write(fila, col, f_antes.strftime('%d-%m-%Y'))
                col += 1
                sheet.write(fila, col, f_antes.strftime('%H:%M:%S'))
                col += 1
                sheet.write(fila, col, f_ahora.strftime('%H:%M:%S'))
                col += 1
                sheet.write(fila, col, horas_)
                trabajo = int(horas_)
                if trabajo == 9:
                    trabajo = 8
                tiempo_no_ocho = horas_ - 9
                if 0 <= tiempo_no_ocho >= TIEMPO_NO_EXTRA:
                    trabajo = horas_
                if tiempo_no_ocho < 0:
                    tiempo_no_ocho = horas_ - 8
                    if 0 <= tiempo_no_ocho >= TIEMPO_NO_EXTRA:
                        trabajo = horas_
                trabajo = round(trabajo, 2)
                col += 1
                sheet.write(fila, col, trabajo)
                nocturna = 0
                extraordinaria = 0
                suplementaria = 0
                if f_antes.weekday() in [calendar.SATURDAY, calendar.SUNDAY] or horas.filtered('festivo'):
                    extraordinaria = trabajo

                jornada = self.env.ref('ecuaminerales_addons_itierp.resource_ocho_horas_1_almuerzo')
                f_fin = f_ahora
                if horas[0].resource_calendar_id == jornada:
                    f_aux = f_antes.replace(hour=17, minute=0, second=0)
                else:
                    f_aux = f_antes.replace(hour=14, minute=0, second=0)
                    if extraordinaria:
                        extraordinaria = extraordinaria - 0.33
                horas_n = (f_fin - f_aux).total_seconds() / 60 / 60
                if 0 < horas_n >= TIEMPO_NO_EXTRA:
                    if trabajo > 8 and not extraordinaria:
                        suplementaria = round(horas_n, 2)
                col += 1
                sheet.write(fila, col, round(nocturna, 2))
                col += 1
                sheet.write(fila, col, round(suplementaria, 2))
                col += 1
                sheet.write(fila, col, round(extraordinaria, 2))
                fila += 1

        return self.return_exel_report(fp, workbook)

    @api.multi
    def print_excel_report_resumen(self):
        fp = BytesIO()
        workbook = xlsxwriter.Workbook(fp)
        sheet = workbook.add_worksheet('RESUMEN DE HORAS')
        format_center = workbook.add_format({'bold': True, 'align': 'vcenter'})
        col = 0
        sheet.set_column(col, col, 50)
        sheet.write(0, col, "EMPLEADO", format_center)
        col += 1
        sheet.set_column(col, col, 20)
        sheet.write(0, col, "JORNADA", format_center)
        col += 1
        sheet.set_column(col, col, 14)
        sheet.write(0, col, "HORAS", format_center)
        col += 1
        sheet.set_column(col, col, 14)
        sheet.write(0, col, "NOCTURNAS", format_center)
        col += 1
        sheet.set_column(col, col, 16)
        sheet.write(0, col, "SUPLEMENTARIAS", format_center)
        col += 1
        sheet.set_column(col, col, 17)
        sheet.write(0, col, "EXTRAORDINARIAS", format_center)
        if not self.hour_production_ids:
            return True

        self.print_excel_report_resumen_del()
        wb = xlrd.open_workbook(file_contents=base64.decodestring(self.file))
        sheet_1 = wb.sheets()[0] if wb.sheets() else None
        sheet_2 = wb.sheets()[1] if wb.sheets() else None
        data_1 = [[sheet_1.cell_value(r, c) for c in range(sheet_1.ncols)] for r in range(sheet_1.nrows)]
        data_2 = [[sheet_2.cell_value(r, c) for c in range(sheet_2.ncols)] for r in range(sheet_2.nrows)]

        fila = 1
        list_data = self.hour_production_ids.filtered(lambda x: x.type_mar in ['exit', 'income'])
        for employee_id in list_data.mapped('employee_id').sorted('name'):
            col = 0
            sheet.write(fila, col, employee_id.display_name)
            col += 1
            sheet.write(fila, col, employee_id.resource_calendar_id.display_name)
            col += 1
            data = list(filter(lambda x: x[0] == employee_id.display_name, data_1))
            if not data:
                data = list(filter(lambda x: x[0] == employee_id.display_name, data_2))

            total_horas_m = 0
            total_nocturnas = 0
            total_extraordinarias = 0
            total_suple = 0
            if not data:
                continue
            for d in data:
                if len(d) == 10:
                    total_horas_m += d[6]
                    total_nocturnas += d[7]
                    total_suple += d[8]
                    total_extraordinarias += d[9]
                else:
                    total_horas_m += d[5]
                    total_nocturnas += d[6]
                    total_suple += d[7]
                    total_extraordinarias += d[8]

            sheet.write(fila, col, total_horas_m)
            col += 1
            sheet.write(fila, col, total_nocturnas)
            col += 1
            sheet.write(fila, col, total_suple)
            col += 1
            sheet.write(fila, col, total_extraordinarias)
            col += 1
            fila += 1

        return self.return_exel_report(fp, workbook)

    def return_exel_report(self, fp, workbook):
        workbook.close()
        self.file = base64.encodestring(fp.getvalue())
        fp.close()
        name_report = "ReporteRoles"
        name_report += '%2Exlsx'
        return {
            'type': 'ir.actions.act_url', 'target': 'new',
            'name': 'contract',
            'url': '/web/content/%s/%s/file/%s?download=true' % (self._name, self.id, name_report),
        }


class FestivosCalculo(models.Model):
    _name = 'production.work.hour.festivo'
    _description = 'Dias festivos calculo'

    parent_id = fields.Many2one('production.work.hour', string='Cabecera')
    festivo_start = fields.Datetime('Inicio Feriado')
    festivo_end = fields.Datetime('Fin Feriado')
