# -*- coding: utf-8 -*-
import re
import base64
import logging
from time import sleep

from tempfile import NamedTemporaryFile

from openerp import api, fields, models, sql_db
from openerp.exceptions import Warning as UserError

from query_helper import QueryHelper


logger = logging.getLogger(__name__)

DEFAULT_ISOLATION_LEVEL = 'READ UNCOMMITTED'
FILENAME_REPLACE_CHARS = [
    (' *', ''),
    (' ', '_'),
    ('*', ''),
    (':', '_'),
]
EXCEL_TITLE_REPLACE_CHARS = [
    (' *', ''),
    ('^', ' '),
    ('&', ' '),
    ('*', ''),
    ('/', ' '),
]


class report_export_base(models.AbstractModel):
    _name = "ibs.report.export.base"
    _description = u"""Base report"""

    xls_file = fields.Binary(u"Файл результатов")
    file_name = fields.Char(string="Filename")

    @api.model
    def initialize(self):
        self.info = self.get_report_info()

    @staticmethod
    def get_workbook():
        """
        Return xlsxwriter.Workbook: workbook

        Raises:
         * UserError: Missing xlsxwriter module
         * UserError: Exception
        """
        try:
            import xlsxwriter
            out_file = NamedTemporaryFile(suffix=".xlsx", delete=True)
            out_file.close()
            workbook = xlsxwriter.Workbook(out_file.name, {"constant_memory": True})
            return workbook, out_file
        except ImportError:
            raise UserError("Missing xlsxwriter module!")
        except Exception as e:
            raise UserError(e)

    @staticmethod
    def get_header_format():
        """
        Return dict: header cell format
        """
        return {
            'bold': True,
            'font_size': 11,
            'fg_color': "#f2f2f2",
            'align': "center",
            'valign': "vcenter",
            'border': True,
            'text_wrap': True,
        }

    @api.model
    def get_title(self):
        """
        Return string: report title
        """
        return self.info.get('title', 'Report title')

    @api.model
    def get_full_filename(self):
        """
        Return string: report full file name
        """

        return u"{}.xlsx".format(
            self.get_filename()
        )

    @staticmethod
    def replace_chars(in_str, chars_list):
        """
        Return string: replace chars in string

        Args:
         * in_str - input string
         * chars_list - list of tuples (string, replace_string)
        """
        for old_char, new_char in chars_list:
            in_str = in_str.replace(old_char, new_char)
        return in_str.strip()

    @api.multi
    def get_filename(self):
        """
        Return string: report file name

        Extra Info:
         * Expected Singleton
        """
        self.ensure_one()
        filename = self.get_title()
        filename = report_export_excel_base.replace_chars(filename, FILENAME_REPLACE_CHARS)
        if getattr(self, 'date_start', None):
            filename += "_{}".format(self.date_start)
        if getattr(self, 'date_end', None):
            filename += "_{}".format(self.date_end)

        return filename

    @api.model
    def get_view_xml_id(self):
        """
        Return string: report view xml id
        """
        return self.info.get('view_id', 'module.view_id')

    @api.model
    def get_action(self, name, xml_id):
        """
        Return dict: view action

        Args:
         * name - action name
         * xml_id - view xml id
        """
        return {
            'name': name,
            'view_type': "form",
            'view_mode': "form",
            'view_id': self.env.ref(xml_id).id,
            'res_model': self._name,
            'context': '{}',
            'type': "ir.actions.act_window",
            'nodestroy': True,
            'target': "new",
            'res_id': self.id,
        }

    @api.model
    def get_download_action(self):
        """
        Return dict: view download action
        """
        return self.get_action(
            self.get_title(),
            self.get_view_xml_id(),
        )

    @api.model
    def get_no_action(self):
        """
        Return dict: view no action
        """
        return self.get_action(
            u"""
            По запросу с данными параметрами ничего не найдено.<br>
            Попробуйте изменить параметры запроса.'
            """,
            "ibs_report_new.ibs_report_new_not_found_view",
        )

    @api.multi
    def get_report_info(self):
        """
        Return dict: report info

        Extra Info:
         * Expected Singleton
        """
        self.ensure_one()
        return {
            'title': "Report title",
            'view_id': "module.view_id",
            'isolation_level': DEFAULT_ISOLATION_LEVEL,
        }

    @api.model
    def write_header(self, workbook, worksheet, data):
        """
        Write report header

        Args:
         * workbook - workbook
         * worksheet - worksheet
         * data - report data

        Returns:
         * int - header rows count
        """
        cell_format = workbook.add_format(self.get_header_format())
        count = 0
        try:
            for col_num, column in enumerate(self.new_env.cr.description):
                worksheet.write(
                    count,
                    col_num,
                    column.name.decode('utf-8'),  # cuz type of name is str
                    cell_format,
                )
            count += 1
        except IndexError:
            pass
        except Exception as e:
            pass

        return count

    @api.model
    def get_params(self):
        """
        Return dict: params for query
        """
        return {}

    @api.multi
    def get_query(self, params):
        """
        Return string or QueryHelper: report query

        Args:
         * params - dict of query params
        """
        return ''

    @api.model
    def get_isolation_level(self):
        """
        Return string: name of isolation level
        """
        return "SET TRANSACTION ISOLATION LEVEL {};\n".format(
            self.info.get('isolation_level', DEFAULT_ISOLATION_LEVEL),
        )

    @api.model
    def construct_query(self):
        """
        Return string: final report query with isolation level as string
        """
        qry = self.get_isolation_level()
        tmp = self.get_query(self.get_params())
        if isinstance(tmp, QueryHelper):
            qry += tmp.query
            qry = tmp.generate_query(qry)
        else:
            qry += tmp

        return qry

    @api.model
    def get_report_data(self):
        """
        Return list of dict: all the remaining rows of a query result set
        """
        query = self.construct_query()
        logger.debug(query)
        new_cr = sql_db.db_connect(self.env.cr.dbname).cursor()
        with api.Environment.manage():
            self.new_env = api.Environment(new_cr, self.env.uid, self.env.context)
            try:
                self.new_env.cr.execute(query)
                res = self.new_env.cr.dictfetchall()
            except Exception as e:
                res = []

        return res

...
