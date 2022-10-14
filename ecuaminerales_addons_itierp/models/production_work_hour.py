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


class ProductionWorkHour(models.Model):
    _name = 'production.work.hour'
    _description = 'Modulo de Horas en producción '
    _rec_name = 'sequence'

    sequence = fields.Char('Secuencia', required=False, readonly=True, track_visibility='onchange')
    document = fields.Binary('Documento de Horas', store=True)
    file_name = fields.Char('File Name', track_visibility='onchange')
    state = fields.Selection([('draft', 'Borrador'),
                              ('load', 'Cargado'),
                              ('purify', 'Depurado'),
                              ('confirm', 'Confirmar'),
                              ('posted', 'Pagado')],
                             readonly=True, default='draft', store=True, string="Estado", track_visibility='onchange')
    search_selection = fields.Selection([('code', 'Codigo Relog'), ('name', 'Nombre')],
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
    file = fields.Binary('document')

    def _compute_count_registers(self):
        self.register_count = len(self.hour_production_ids)

    def view_registro_horas(self):
        self.ensure_one()
        return {
            'name': 'Registro de Horas',
            'type': 'ir.actions.act_window',
            'view_mode': 'tree,form',
            'res_model': 'production.work.hour.employee',
            'domain': [('id', 'in', self.hour_production_ids.ids)],
        }

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
        if data[0] != ['Nombre', 'Número de empleado', 'Departamento', 'Fecha', 'Hora', 'Dispositivo']:
            raise ValidationError("""Recurde que el archivo debe contener la siguiente estructura  \n
            ['Nombre', 'Número de empleado', 'Departamento', 'Fecha', 'Hora', 'Dispositivo']""")
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

    def detectar_ingreso_salida(self, antes, ahora, minutes):
        # TURNO ROTATIVOS
        sales_journal_id = self.env.ref('ecuaminerales_addons_itierp.resource_rotativos')
        if ahora.resource_calendar_id == sales_journal_id:
            # VALIDAR HORARIOS
            # 	             |      TURNO 1	   | TURNO 2	        | TURNO 3
            # LUNES-VIERNES	 | 06H00 - 14H00   |	14H00 - 22H00	| 22H00 - 06H00
            # SÁBADO	     |06H00 - 18H00	   |    LIBRE	        | 18H00 - 06H00
            # DOMINGO	     |06H00 - 18H00    |	18H00 - 06H00	| LIBRE
            f_antes = antes.fecha_time - timedelta(hours=5)
            f_ahora = ahora.fecha_time - timedelta(hours=5)
            if f_antes.day == 26:
                print("To ca ver que pasa aca")
            if antes.turno != 'no':
                return True
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
                if 5 <= f_antes.hour <= 7 and f_ahora.hour <= 20:
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
            print("Toca ver por que llegas hasta aca....")

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
        for employee_id in set(data_filter.mapped('employee_id')):
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

    def delete_duplicates(self):
        self.hour_production_ids = self.hour_production_ids.filtered(lambda x: not x.delete)
        self.purge_data()
        self.state = 'purify'
        self.turnos_rotativos_html_insertion()

    def data_turnos_of_day(self, marcaciones):
        if not marcaciones:
            return ["", "", "", ""]
        data = ["", "", "", ""]
        # TURNO 1
        t1 = marcaciones.filtered(lambda x: x.turno == 't1')
        if t1:
            fechas = max(t1.mapped('fecha_time')) - min(t1.mapped('fecha_time'))
            data[0] = fechas.total_seconds() / 60 / 60
        t2 = marcaciones.filtered(lambda x: x.turno == 't2')
        if t2:
            fechas = max(t2.mapped('fecha_time')) - min(t2.mapped('fecha_time'))
            data[1] = fechas.total_seconds() / 60 / 60
        t3 = marcaciones.filtered(lambda x: x.turno == 't3')
        if t3:
            fechas = max(t3.mapped('fecha_time')) - min(t3.mapped('fecha_time'))
            data[2] = fechas.total_seconds() / 60 / 60
        tt2 = marcaciones.filtered(lambda x: x.turno == 'tt2')
        if tt2:
            fechas = max(tt2.mapped('fecha_time')) - min(tt2.mapped('fecha_time'))
            data[3] = fechas.total_seconds() / 60 / 60
        t1f = marcaciones.filtered(lambda x: x.turno == 't1f')
        if t1f:
            fechas = max(t1f.mapped('fecha_time')) - min(t1f.mapped('fecha_time'))
            data[0] = fechas.total_seconds() / 60 / 60
        t2f = marcaciones.filtered(lambda x: x.turno == 't2f')
        if t2f:
            fechas = max(t2f.mapped('fecha_time')) - min(t2f.mapped('fecha_time'))
            data[1] = fechas.total_seconds() / 60 / 60
        t3f = marcaciones.filtered(lambda x: x.turno == 't3f')
        if t3f:
            fechas = max(t3f.mapped('fecha_time')) - min(t3f.mapped('fecha_time'))
            data[2] = fechas.total_seconds() / 60 / 60
        return data

    def print_header_excel(self, sheet, format_center):
        sheet.set_column(0, 0, 45)
        sheet.merge_range(0, 0, 1, 0, "Empleado", format_center)
        fecha_header = self.fecha_inicio.strftime('%d-%m')
        count = 1
        for day in range(self.number_of_days):
            sheet.merge_range(0, count, 0, count + 3, fecha_header, format_center)
            sheet.set_column(count, count + 3, 3)
            sheet.write(1, count, "T1", format_center)
            sheet.write(1, count + 1, "T2", format_center)
            sheet.write(1, count + 2, "T3", format_center)
            sheet.write(1, count + 3, "TD", format_center)
            fecha_header = (self.fecha_inicio + timedelta(days=day + 1)).strftime('%d-%m')
            count += 4

    def excel_turnos_rotativos(self, sheet, format_center):
        sales_journal_id = self.env.ref('ecuaminerales_addons_itierp.resource_rotativos')
        data_filter = self.hour_production_ids.filtered(
            lambda x: x.resource_calendar_id == sales_journal_id and x.turno != 'no')
        fila = 2
        for employee_id in set(data_filter.mapped('employee_id')):
            sheet.write(fila, 0, employee_id.display_name, format_center)
            list_hours = data_filter.filtered(lambda x: x.employee_id == employee_id).sorted('fecha_time')
            fecha_header = self.fecha_inicio - timedelta(hours=5)
            col = 1
            for day in range(1, self.number_of_days + 1):
                data = list_hours.filtered(
                    lambda x: x.fecha_time.strftime('%d-%m-%y') == fecha_header.strftime('%d-%m-%y') and x.turno in [
                        't1', 't1f'])
                self.print_data_lina(data, col, fila, sheet)
                col += 1
                data = list_hours.filtered(
                    lambda x: x.fecha_time.strftime('%d-%m-%y') == fecha_header.strftime('%d-%m-%y') and x.turno in [
                        't2'])
                self.print_data_lina(data, col, fila, sheet)
                col += 1
                fecha_nex = self.fecha_inicio + timedelta(days=day) - timedelta(hours=5)

                data = list_hours.filtered(
                    lambda x: fecha_header < x.fecha_time < fecha_nex and x.turno in ['t3', 't2f', 't3f'])
                if len(data) > 2:
                    print("Herer")
                self.print_data_lina(data, col, fila, sheet)
                col += 1
                sheet.write(fila, col, "X")
                col += 1

                fecha_header = self.fecha_inicio + timedelta(days=day) - timedelta(hours=5)

            fila += 1

    def print_data_lina(self, data, col, fila, sheet):
        if data:
            data = data.sorted('fecha_time')
            if not len(data) > 1:
                sheet.write(fila, col, 1)
            else:
                horas = data[1].fecha_time - data[0].fecha_time
                horas = round(horas.total_seconds() / 60 / 60, 2)
                sheet.write(fila, col, horas)
        else:
            sheet.write(fila, col, '')

    @api.multi
    def print_excel_report(self):
        fp = BytesIO()
        workbook = xlsxwriter.Workbook(fp)
        sheet = workbook.add_worksheet('Turnos Rotativos')
        format_center = workbook.add_format({'bold': True, 'align': 'vcenter'})
        self.print_header_excel(sheet, format_center)
        self.excel_turnos_rotativos(sheet, format_center)
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
