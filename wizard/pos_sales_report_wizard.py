import io
import base64
from datetime import datetime, time

import pytz

from odoo import api, fields, models, _
from odoo.exceptions import UserError


class PosSalesReportWizard(models.TransientModel):
    _name = 'pos.sales.report.wizard'
    _description = 'Wizard Rapport Ventes PdV'

    date_from = fields.Date(
        string='Date debut', required=True,
        default=lambda self: fields.Date.today().replace(day=1),
    )
    date_to = fields.Date(
        string='Date fin', required=True,
        default=fields.Date.today,
    )
    pos_config_ids = fields.Many2many(
        'pos.config', string='Point(s) de vente',
        help='Laisser vide pour tous les PdV.',
    )
    product_ids = fields.Many2many(
        'product.product', string='Produits',
        help='Laisser vide pour tous les produits.',
    )
    categ_ids = fields.Many2many(
        'product.category', string='Familles produit',
        help='Laisser vide pour toutes les familles.',
    )
    show_pos_summary = fields.Boolean(
        string='Synthèse CA par PdV',
        help='Ajoute un onglet avec le CA détaillé par point de vente.',
    )
    report_file = fields.Binary('Fichier', readonly=True)
    report_filename = fields.Char('Nom du fichier', readonly=True)

    def _get_pos_summary(self, rows):
        """Regroupe les données par PdV pour la synthèse CA."""
        summary = {}
        for r in rows:
            pos_name = r['pos_name'] or 'Non défini'
            if pos_name not in summary:
                summary[pos_name] = {'pos_name': pos_name, 'ca_ht': 0.0, 'ca_ttc': 0.0, 'nb_lignes': 0, 'total_qty': 0.0}
            summary[pos_name]['ca_ht'] += r['mtt_ht']
            summary[pos_name]['ca_ttc'] += r['mtt_ttc']
            summary[pos_name]['total_qty'] += r['qty']
            summary[pos_name]['nb_lignes'] += 1
        result = list(summary.values())
        result.sort(key=lambda x: x['ca_ttc'], reverse=True)
        return result

    def action_export_excel(self):
        self.ensure_one()
        if self.date_from > self.date_to:
            raise UserError(_('La date de debut doit etre anterieure a la date de fin.'))

        rows = self._get_data()
        pos_summary = self._get_pos_summary(rows) if self.show_pos_summary else None
        content = self._generate_xlsx(rows, pos_summary=pos_summary)
        self.report_file = base64.b64encode(content)
        self.report_filename = 'rapport_pdv_%s_%s.xlsx' % (
            self.date_from.strftime('%Y%m%d'),
            self.date_to.strftime('%Y%m%d'),
        )
        return {
            'type': 'ir.actions.act_url',
            'url': '/web/content/?model=%s&id=%d&field=report_file'
                   '&filename_field=report_filename&download=true' % (
                       self._name, self.id),
            'target': 'new',
        }

    def _get_data(self):
        tz = pytz.timezone(self.env.user.tz or 'Indian/Antananarivo')
        date_from_dt = tz.localize(datetime.combine(
            self.date_from, time.min,
        )).astimezone(pytz.utc).replace(tzinfo=None)
        date_to_dt = tz.localize(datetime.combine(
            self.date_to, time.max,
        )).astimezone(pytz.utc).replace(tzinfo=None)

        domain = [
            ('order_id.date_order', '>=', date_from_dt),
            ('order_id.date_order', '<=', date_to_dt),
        ]
        if self.pos_config_ids:
            domain.append(('order_id.config_id', 'in', self.pos_config_ids.ids))
        if self.product_ids:
            domain.append(('product_id', 'in', self.product_ids.ids))
        if self.categ_ids:
            domain.append(('product_id.categ_id', 'in', self.categ_ids.ids))

        lines = self.env['pos.order.line'].search(
            domain, order='order_id, id',
        )

        user_tz = pytz.timezone(self.env.user.tz or 'Indian/Antananarivo')
        rows = []
        for line in lines:
            order = line.order_id
            product = line.product_id
            categ = product.categ_id
            dt = order.date_order
            if dt and dt.tzinfo is None:
                dt = pytz.utc.localize(dt)
            local_dt = dt.astimezone(user_tz) if dt else False

            rows.append({
                'date': local_dt.strftime('%d/%m/%Y') if local_dt else '',
                'code_article': product.default_code or '',
                'code_famille': categ.complete_name.split(' / ')[-1] if categ else '',
                'famille': categ.name if categ else '',
                'designation': product.name or '',
                'pv': line.price_unit or 0.0,
                'pos_name': order.config_id.name if order.config_id else '',
                'qty': line.qty or 0.0,
                'mtt_ht': line.price_subtotal or 0.0,
                'mtt_ttc': line.price_subtotal_incl or 0.0,
            })
        return rows

    def _generate_xlsx(self, rows, pos_summary=None):
        import xlsxwriter

        output = io.BytesIO()
        wb = xlsxwriter.Workbook(output, {'in_memory': True})

        fmt_title = wb.add_format({
            'bold': True, 'font_size': 14, 'align': 'center',
        })
        fmt_header = wb.add_format({
            'bold': True, 'bg_color': '#4472C4', 'font_color': 'white',
            'border': 1, 'align': 'center', 'text_wrap': True,
        })
        fmt_header_green = wb.add_format({
            'bold': True, 'bg_color': '#00FF00', 'font_color': '#000080',
            'border': 1, 'align': 'center', 'text_wrap': True,
        })
        fmt_text = wb.add_format({'border': 1, 'font_size': 10})
        fmt_num = wb.add_format({
            'border': 1, 'font_size': 10, 'num_format': '#,##0.00',
        })
        fmt_qty = wb.add_format({
            'border': 1, 'font_size': 10, 'num_format': '#,##0.000',
        })
        fmt_qty_neg = wb.add_format({
            'border': 1, 'font_size': 10, 'num_format': '#,##0.000',
            'font_color': 'red',
        })
        fmt_num_neg = wb.add_format({
            'border': 1, 'font_size': 10, 'num_format': '#,##0.00',
            'font_color': 'red',
        })
        fmt_total_lbl = wb.add_format({
            'bold': True, 'bg_color': '#1F3864', 'font_color': 'white',
            'border': 2, 'font_size': 11,
        })
        fmt_total_num = wb.add_format({
            'bold': True, 'bg_color': '#1F3864', 'font_color': 'white',
            'border': 2, 'font_size': 11, 'num_format': '#,##0.00',
        })
        fmt_total_qty = wb.add_format({
            'bold': True, 'bg_color': '#1F3864', 'font_color': 'white',
            'border': 2, 'font_size': 11, 'num_format': '#,##0.000',
        })

        ws = wb.add_worksheet('Analyse Ventes PdV')

        headers = [
            ('DATE', fmt_header, 14),
            ('CODE ARTICLES', fmt_header, 16),
            ('Code famille', fmt_header, 16),
            ('Famille', fmt_header_green, 20),
            ('DESIGNATION ARTICLES', fmt_header, 36),
            ('PV', fmt_header, 12),
            ('Point de vente', fmt_header, 22),
            ('QTES', fmt_header, 12),
            ('MTT HT', fmt_header, 14),
            ('MTT TTC', fmt_header, 14),
        ]

        for i, (_, _, w) in enumerate(headers):
            ws.set_column(i, i, w)

        last_col = len(headers) - 1
        ws.merge_range(0, 0, 0, last_col, 'Analyse Ventes Point de Vente', fmt_title)
        ws.write(1, 0, self.env.company.name, fmt_text)
        ws.write(1, last_col - 1, 'Periode', fmt_text)
        ws.write(1, last_col, '%s au %s' % (
            self.date_from.strftime('%d/%m/%Y'),
            self.date_to.strftime('%d/%m/%Y'),
        ), fmt_text)

        row = 3
        for col, (label, fmt, _) in enumerate(headers):
            ws.write(row, col, label, fmt)
        row += 1
        ws.freeze_panes(row, 0)
        first_data = row

        total_qty = 0.0
        total_ht = 0.0
        total_ttc = 0.0

        for r in rows:
            ws.write(row, 0, r['date'], fmt_text)
            ws.write(row, 1, r['code_article'], fmt_text)
            ws.write(row, 2, r['code_famille'], fmt_text)
            ws.write(row, 3, r['famille'], fmt_text)
            ws.write(row, 4, r['designation'], fmt_text)
            ws.write(row, 5, r['pv'], fmt_num if r['pv'] >= 0 else fmt_num_neg)
            ws.write(row, 6, r['pos_name'], fmt_text)
            ws.write(row, 7, r['qty'], fmt_qty if r['qty'] >= 0 else fmt_qty_neg)
            ws.write(row, 8, r['mtt_ht'], fmt_num if r['mtt_ht'] >= 0 else fmt_num_neg)
            ws.write(row, 9, r['mtt_ttc'], fmt_num if r['mtt_ttc'] >= 0 else fmt_num_neg)
            total_qty += r['qty']
            total_ht += r['mtt_ht']
            total_ttc += r['mtt_ttc']
            row += 1

        if row > first_data:
            ws.autofilter(first_data - 1, 0, row - 1, last_col)

        ws.merge_range(row, 0, row, 6, 'TOTAL', fmt_total_lbl)
        ws.write(row, 7, total_qty, fmt_total_qty)
        ws.write(row, 8, total_ht, fmt_total_num)
        ws.write(row, 9, total_ttc, fmt_total_num)

        # Onglet Synthèse CA par PdV
        if pos_summary:
            ws2 = wb.add_worksheet('CA par PdV')
            ws2.merge_range(0, 0, 0, 4,
                            'Synthèse CA par Point de Vente', fmt_title)
            ws2.write(1, 0, self.env.company.name, fmt_text)
            ws2.write(1, 3, 'Période', fmt_text)
            ws2.write(1, 4, '%s au %s' % (
                self.date_from.strftime('%d/%m/%Y'),
                self.date_to.strftime('%d/%m/%Y'),
            ), fmt_text)

            s_headers = [
                ('POINT DE VENTE', fmt_header, 30),
                ('NB LIGNES', fmt_header, 12),
                ('QTÉ TOTALE', fmt_header, 14),
                ('CA HT', fmt_header, 16),
                ('CA TTC', fmt_header, 16),
            ]
            for i, (_, _, w) in enumerate(s_headers):
                ws2.set_column(i, i, w)

            s_row = 3
            for col, (label, fmt, _) in enumerate(s_headers):
                ws2.write(s_row, col, label, fmt)
            s_row += 1
            ws2.freeze_panes(s_row, 0)

            s_total_lines = 0
            s_total_qty = 0.0
            s_total_ht = 0.0
            s_total_ttc = 0.0

            for r in pos_summary:
                ws2.write(s_row, 0, r['pos_name'], fmt_text)
                ws2.write(s_row, 1, r['nb_lignes'], fmt_num)
                ws2.write(s_row, 2, r['total_qty'], fmt_qty if r['total_qty'] >= 0 else fmt_qty_neg)
                ws2.write(s_row, 3, r['ca_ht'], fmt_num if r['ca_ht'] >= 0 else fmt_num_neg)
                ws2.write(s_row, 4, r['ca_ttc'], fmt_num if r['ca_ttc'] >= 0 else fmt_num_neg)
                s_total_lines += r['nb_lignes']
                s_total_qty += r['total_qty']
                s_total_ht += r['ca_ht']
                s_total_ttc += r['ca_ttc']
                s_row += 1

            ws2.autofilter(3, 0, s_row - 1, 4)
            ws2.write(s_row, 0, 'TOTAL', fmt_total_lbl)
            ws2.write(s_row, 1, s_total_lines, fmt_total_num)
            ws2.write(s_row, 2, s_total_qty, fmt_total_qty)
            ws2.write(s_row, 3, s_total_ht, fmt_total_num)
            ws2.write(s_row, 4, s_total_ttc, fmt_total_num)

        wb.close()
        return output.getvalue()
