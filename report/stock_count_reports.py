from odoo import models
import logging
import xlwt
_logger = logging.getLogger(__name__)
class PartnerXlsx(models.AbstractModel):
    _name = 'report.stock_count_report.stock_count_report_xlsx'
    _inherit = 'report.report_xlsx.abstract'
    def generate_xlsx_report(self, workbook, data, partners):
        _logger.info('Generating XLS report')
        _logger.info('Generating XLS report')
   
        _logger.info(data)
        sheet = workbook.add_worksheet('Stock Count Report')
        bold = workbook.add_format({'bold': True})
        sheet.write(0, 0, 'Product Name', bold)
        sheet.write(0, 1, 'Location', bold)
        sheet.write(0, 2, 'Lot/Serial Number', bold)
        sheet.write(0, 3, 'On Hand Quantity', bold)
        sheet.write(0, 4, 'Reserved Quantity', bold)
        sheet.write(0, 5, 'Price', bold)

        stock_count_data = data['stock_count_data']
        row = 1
        for stock_count in stock_count_data:
            sheet.write(row, 0, stock_count['name'])
            sheet.write(row, 1, stock_count['location'])
            sheet.write(row, 2, stock_count['lot'])
            sheet.write(row, 3, stock_count['inventory_quantity'])
            sheet.write(row, 4, stock_count['reserved_quantity'])
            sheet.write(row, 5, stock_count['value'])
            row += 1


       