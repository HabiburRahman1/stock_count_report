import base64
from io import BytesIO
import logging
import xlwt
from odoo import fields, models, api, _
import asyncio
_logger = logging.getLogger(__name__)


class StockCountReportWizard(models.TransientModel):
    _name = 'stock.count.report.wizard'

    start_date = fields.Date(
        string='Start Date', required=True, default=fields.Date.context_today)
    end_date = fields.Date(string='End Date', required=True,
                           default=fields.Date.context_today)

    def print_report(self):
        result_stock_count = []

        async def get_stock_count_data(task_no):
            page_number = task_no
            per_page = 100
            offset = 0 if page_number <= 1 else (page_number - 1) * per_page
            stock_object = self.env['stock.quant'].search(
                [('location_id.usage', '=', 'internal')], order="id asc", offset=offset, limit=per_page
            )
            data = []
            for stock in stock_object:

                data.append({
                    'x_product_default_code': stock.x_product_default_code,
                    'name': stock.product_id.name,
                    'location': stock.location_id.name,
                    'lot': stock.lot_id.name,
                    'product_manufacturer_id': stock.product_manufacturer_id.name,
                    'category_product_id': stock.category_product_id.name,
                    'location_warehouse_id': stock.location_warehouse_id.name,
                    'sale_id': stock.sale_id.name,
                    'so_amount_total': stock.so_amount_total,
                    'sale_date_order': stock.sale_date_order,
                    'po_name': stock.po_name,
                    'po_amount_total': stock.po_amount_total,
                    'purchase_date_order': stock.purchase_date_order,
                    'unit_price': stock.unit_price,
                    'package_id': stock.package_id.name,
                    'inventory_quantity': stock.inventory_quantity,
                    'reserved_quantity': stock.reserved_quantity,
                    'product_uom_id': stock.product_uom_id.name,
                    'value': stock.value,
                    'company_id': stock.company_id.name,
                })
            return data

        async def async_get_stock_count_data():
            data_list = []
            stock_object_count = self.env['stock.quant'].search_count(
                [('location_id.usage', '=', 'internal')])
            page_count = stock_object_count // 100
            for i in range(1, page_count):
                data_list.append(
                    asyncio.create_task(get_stock_count_data(i))
                )
            responses = await asyncio.gather(*data_list)
            for response in responses:
                result_stock_count.extend(response)

        # Async Loop for testing
        try:
            asyncio.run(async_get_stock_count_data())
        except Exception as e:
            pass
        return self.env.ref('stock_count_report.stock_count_report_xlsx').report_action(self, data={"stock_count_data": result_stock_count})
