using System;
using System.Collections.Generic;
using System.Windows;
using ClosedXML.Excel;
using JSON_Tools.Models;

namespace JSON_Tools.Services
{
    public class ExcelExportService
    {
        public void ExportToExcel(object data, string filePath)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Export Daten");

                if (data is List<Json1Order> list1) ExportJson1(worksheet, list1);
                else if (data is List<Json2Order> list2) ExportJson2(worksheet, list2);
                else if (data is List<Json3Order> list3) ExportJson3(worksheet, list3);
                else
                {
                    throw new ArgumentException("Unbekanntes Format für den Export.");
                }

                worksheet.Columns().AdjustToContents();
                workbook.SaveAs(filePath);
            }
        }

        private void ExportJson1(IXLWorksheet worksheet, List<Json1Order> orders)
        {
            worksheet.Cell(1, 1).Value = "ID";
            worksheet.Cell(1, 2).Value = "Kunde";
            worksheet.Cell(1, 3).Value = "Erstellt";
            worksheet.Cell(1, 4).Value = "Gesamtbetrag";
            StyleHeader(worksheet, 4);

            int row = 2;
            foreach (var order in orders)
            {
                worksheet.Cell(row, 1).Value = order.OrderId;
                worksheet.Cell(row, 2).Value = order.Customer;
                worksheet.Cell(row, 3).Value = order.Created;
                worksheet.Cell(row, 3).Style.DateFormat.Format = "dd.MM.yyyy";
                worksheet.Cell(row, 4).Value = order.Amount;
                worksheet.Cell(row, 4).Style.NumberFormat.Format = "#,##0.00 €";
                row++;
            }
            worksheet.Columns().AdjustToContents();
        }

        private void ExportJson2(IXLWorksheet worksheet, List<Json2Order> orders)
        {
            worksheet.Cell(1, 1).Value = "ID";
            worksheet.Cell(1, 2).Value = "Kunde";
            worksheet.Cell(1, 3).Value = "Erstellt";
            worksheet.Cell(1, 4).Value = "SKU";
            worksheet.Cell(1, 5).Value = "Menge";
            worksheet.Cell(1, 6).Value = "Einzelpreis";
            worksheet.Cell(1, 7).Value = "Zeilenwert";
            worksheet.Cell(1, 8).Value = "Gesamtbetrag";
            StyleHeader(worksheet, 8);

            int row = 2;
            foreach (var order in orders)
            {
                if (order.Items == null || order.Items.Count == 0)
                {
                    worksheet.Cell(row, 1).Value = order.OrderId;
                    worksheet.Cell(row, 2).Value = order.Customer;
                    worksheet.Cell(row, 3).Value = order.Created;
                    worksheet.Cell(row, 3).Style.DateFormat.Format = "dd.MM.yyyy";
                    worksheet.Cell(row, 4).Value = "-";
                    worksheet.Cell(row, 5).Value = "-";
                    worksheet.Cell(row, 6).Value = "-";
                    worksheet.Cell(row, 7).Value = 0;
                    worksheet.Cell(row, 8).Value = order.TotalAmount;
                    worksheet.Cell(row, 8).Style.NumberFormat.Format = "#,##0.00 €";
                    row++;
                }
                else
                {
                    foreach (var item in order.Items)
                    {
                        worksheet.Cell(row, 1).Value = order.OrderId;
                        worksheet.Cell(row, 2).Value = order.Customer;
                        worksheet.Cell(row, 3).Value = order.Created;
                        worksheet.Cell(row, 3).Style.DateFormat.Format = "dd.MM.yyyy";
                        worksheet.Cell(row, 4).Value = item.Sku;
                        worksheet.Cell(row, 5).Value = item.Qty;
                        worksheet.Cell(row, 6).Value = item.Price;
                        worksheet.Cell(row, 6).Style.NumberFormat.Format = "#,##0.00 €";
                        worksheet.Cell(row, 7).Value = item.ItemTotal;
                        worksheet.Cell(row, 7).Style.NumberFormat.Format = "#,##0.00 €";
                        worksheet.Cell(row, 8).Value = order.TotalAmount;
                        worksheet.Cell(row, 8).Style.NumberFormat.Format = "#,##0.00 €";
                        row++;
                    }
                }
            }
            worksheet.Columns().AdjustToContents();
        }

        private void ExportJson3(IXLWorksheet worksheet, List<Json3Order> orders)
        {
            worksheet.Cell(1, 1).Value = "ID";
            worksheet.Cell(1, 2).Value = "Kunde";
            worksheet.Cell(1, 3).Value = "Erstellt";
            worksheet.Cell(1, 4).Value = "Status";
            worksheet.Cell(1, 5).Value = "Vertreter";
            worksheet.Cell(1, 6).Value = "SKU";
            worksheet.Cell(1, 7).Value = "Menge";
            worksheet.Cell(1, 8).Value = "Einzelreis";
            worksheet.Cell(1, 9).Value = "Zeilenwert";
            worksheet.Cell(1, 10).Value = "Lieferadresse";
            worksheet.Cell(1, 11).Value = "Lieferdatum";
            worksheet.Cell(1, 12).Value = "Gesamtbetrag";
            StyleHeader(worksheet, 12);

            int row = 2;
            foreach (var order in orders)
            {
                if (order.Items == null || order.Items.Count == 0)
                {
                    worksheet.Cell(row, 1).Value = order.OrderId;
                    worksheet.Cell(row, 2).Value = order.Customer;
                    worksheet.Cell(row, 3).Value = order.Created;
                    worksheet.Cell(row, 3).Style.DateFormat.Format = "dd.MM.yyyy";
                    worksheet.Cell(row, 4).Value = order.Status;
                    worksheet.Cell(row, 5).Value = order.SalesRep;
                    worksheet.Cell(row, 6).Value = "-";
                    worksheet.Cell(row, 7).Value = "-";
                    worksheet.Cell(row, 8).Value = 0;
                    worksheet.Cell(row, 9).Value = 0;
                    worksheet.Cell(row, 10).Value = order.Delivery?.Address;

                    if (order.Delivery?.DeliveryDate != null)
                    {
                        worksheet.Cell(row, 11).Value = order.Delivery.DeliveryDate;
                        worksheet.Cell(row, 3).Style.DateFormat.Format = "dd.MM.yyyy";
                    }
                    else worksheet.Cell(row, 11).Value = "-";

                    worksheet.Cell(row, 12).Value = order.TotalAmount;
                    worksheet.Cell(row, 12).Style.NumberFormat.Format = "#,##0.00 €";
                    row++;
                }

                else
                {
                    foreach (var item in order.Items)
                    {
                        worksheet.Cell(row, 1).Value = order.OrderId;
                        worksheet.Cell(row, 2).Value = order.Customer;
                        worksheet.Cell(row, 3).Value = order.Created;
                        worksheet.Cell(row, 3).Style.DateFormat.Format = "dd.MM.yyyy";
                        worksheet.Cell(row, 4).Value = order.Status;
                        worksheet.Cell(row, 5).Value = order.SalesRep;
                        worksheet.Cell(row, 6).Value = item.Sku;
                        worksheet.Cell(row, 7).Value = item.Qty;
                        worksheet.Cell(row, 8).Value = item.Price;
                        worksheet.Cell(row, 8).Style.NumberFormat.Format = "#,##0.00 €";
                        worksheet.Cell(row, 9).Value = item.ItemTotal;
                        worksheet.Cell(row, 9).Style.NumberFormat.Format = "#,##0.00 €";
                        worksheet.Cell(row, 10).Value = order.Delivery?.Address;

                        if (order.Delivery?.DeliveryDate != null)
                        {
                            worksheet.Cell(row, 11).Value = order.Delivery.DeliveryDate;
                            worksheet.Cell(row, 3).Style.DateFormat.Format = "dd.MM.yyyy";
                        }
                        else worksheet.Cell(row, 11).Value = "-";

                        worksheet.Cell(row, 12).Value = order.TotalAmount;
                        worksheet.Cell(row, 12).Style.NumberFormat.Format = "#,##0.00 €";
                        row++;
                    }
                }
            }
            worksheet.Columns().AdjustToContents();
        }

        private void StyleHeader(IXLWorksheet ws, int colCount)
        {
            var range = ws.Range(1, 1, 1, colCount);
            range.Style.Font.Bold = true;
            range.Style.Fill.BackgroundColor = XLColor.LightGray;
        }
    }
}
