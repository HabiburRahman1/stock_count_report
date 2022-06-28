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
        sheet.write(0, 0, 'Internal Reference', bold)
        sheet.write(0, 1, 'Product', bold)
        sheet.write(0, 2, 'Location', bold)
        sheet.write(0, 3, 'Lot/Serial Number', bold)
        sheet.write(0, 4, 'Menufacturer', bold)
        sheet.write(0, 5, 'Category', bold)
        sheet.write(0, 6, 'Warehouse', bold)
        sheet.write(0, 7, 'Last Order', bold)
        sheet.write(0, 8, 'SO Amount', bold)
        sheet.write(0, 9, 'Date SO', bold)
        sheet.write(0, 10, 'Last Order', bold)
        sheet.write(0, 11, 'PO Amount', bold)
        sheet.write(0, 12, 'Date PO', bold)
        sheet.write(0, 13, 'Unit Price', bold)
        sheet.write(0, 14, 'Package', bold)
        sheet.write(0, 15, 'On Hand Quantity', bold)
        sheet.write(0, 16, 'Reserved Quantity', bold)
        sheet.write(0, 17, 'UOM', bold)
        sheet.write(0, 18, 'Value', bold)
        sheet.write(0, 19, 'Company', bold)

        stock_count_data = data['stock_count_data']
        row = 1
        for stock_count in stock_count_data:
            sheet.write(row, 0, stock_count['x_product_default_code'])
            sheet.write(row, 1, stock_count['name'])
            sheet.write(row, 2, stock_count['location'])
            sheet.write(row, 3, stock_count['lot'])
            sheet.write(row, 4, stock_count['product_manufacturer_id'])
            sheet.write(row, 5, stock_count['category_product_id'])
            sheet.write(row, 6, stock_count['location_warehouse_id'])
            sheet.write(row, 7, stock_count['sale_id'])
            sheet.write(row, 8, stock_count['so_amount_total'])
            sheet.write(row, 9, stock_count['sale_date_order'])
            sheet.write(row, 10, stock_count['po_name'])
            sheet.write(row, 11, stock_count['po_amount_total'])
            sheet.write(row, 12, stock_count['purchase_date_order'])
            sheet.write(row, 13, stock_count['unit_price'])
            sheet.write(row, 14, stock_count['package_id'])
            sheet.write(row, 15, stock_count['inventory_quantity'])
            sheet.write(row, 16, stock_count['reserved_quantity'])
            sheet.write(row, 17, stock_count['product_uom_id'])
            sheet.write(row, 18, stock_count['value'])
            sheet.write(row, 19, stock_count['company_id'])
            row += 1


       