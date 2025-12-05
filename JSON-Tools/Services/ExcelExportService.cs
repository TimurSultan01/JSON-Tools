using System;
using System.Collections.Generic;
using ClosedXML.Excel;
using JSON_Tools.Models;

namespace JSON_Tools.Services
{
    public class ExcelExportService
    {
        public void ExportToExcel(List<Order> orders, string filePath)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Export Daten");

                // Add headers
                worksheet.Cell(1, 1).Value = "ID";
                worksheet.Cell(1, 2).Value = "Kunde";
                worksheet.Cell(1, 3).Value = "Datum";
                worksheet.Cell(1, 4).Value = "Gesamtwert";

                // Format header
                var headerRange = worksheet.Range("A1:D1");
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;

                // Add order data
                int row = 2;
                foreach (var order in orders)
                {
                    worksheet.Cell(row, 1).Value = order.OrderId;
                    worksheet.Cell(row, 2).Value = order.Customer;
                    worksheet.Cell(row, 3).Value = order.Created;
                    worksheet.Cell(row, 3).Style.DateFormat.Format = "yyyy.MM.dd";
                    worksheet.Cell(row, 4).Value = order.Amount;
                    worksheet.Cell(row, 4).Style.NumberFormat.Format = "#,##0.00 €";

                    row++;
                }

                worksheet.Columns().AdjustToContents();
                workbook.SaveAs(filePath);
            }
        }
    }
}
