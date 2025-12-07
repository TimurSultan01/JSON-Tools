using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using JSON_Tools.Models;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;

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
            // After saving with ClosedXML, add charts with EPPlus
            AddChartsToExcel(filePath, data);
        }

        // Helper method to handle null or empty values
        private string GetVal(string value)
        {
            return string.IsNullOrEmpty(value) ? "-" : value;
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
                worksheet.Cell(row, 2).Value = GetVal(order.Customer);
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
                    worksheet.Cell(row, 2).Value = GetVal(order.Customer);
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
                        worksheet.Cell(row, 2).Value = GetVal(order.Customer);
                        worksheet.Cell(row, 3).Value = order.Created;
                        worksheet.Cell(row, 3).Style.DateFormat.Format = "dd.MM.yyyy";
                        worksheet.Cell(row, 4).Value = GetVal(item.Sku);
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
                    worksheet.Cell(row, 2).Value = GetVal(order.Customer);
                    worksheet.Cell(row, 3).Value = order.Created;
                    worksheet.Cell(row, 3).Style.DateFormat.Format = "dd.MM.yyyy";
                    worksheet.Cell(row, 4).Value = GetVal(order.Status);
                    worksheet.Cell(row, 5).Value = GetVal(order.SalesRep);
                    worksheet.Cell(row, 6).Value = "-";
                    worksheet.Cell(row, 7).Value = "-";
                    worksheet.Cell(row, 8).Value = 0;
                    worksheet.Cell(row, 9).Value = 0;
                    worksheet.Cell(row, 10).Value = GetVal(order.Delivery?.Address);

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
                        worksheet.Cell(row, 2).Value = GetVal(order.Customer);
                        worksheet.Cell(row, 3).Value = order.Created;
                        worksheet.Cell(row, 3).Style.DateFormat.Format = "dd.MM.yyyy";
                        worksheet.Cell(row, 4).Value = GetVal(order.Status);
                        worksheet.Cell(row, 5).Value = GetVal(order.SalesRep);
                        worksheet.Cell(row, 6).Value = GetVal(item.Sku);
                        worksheet.Cell(row, 7).Value = item.Qty;
                        worksheet.Cell(row, 8).Value = item.Price;
                        worksheet.Cell(row, 8).Style.NumberFormat.Format = "#,##0.00 €";
                        worksheet.Cell(row, 9).Value = item.ItemTotal;
                        worksheet.Cell(row, 9).Style.NumberFormat.Format = "#,##0.00 €";
                        worksheet.Cell(row, 10).Value = GetVal(order.Delivery?.Address);

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

        // Method to add charts using EPPlus
        private void AddChartsToExcel(string filePath, object data)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var chartSheet = package.Workbook.Worksheets.Add("Dashboard & Analyse");

                // Make charts based on data type
                if (data is List<Json1Order> list1) AddChartsJson1(chartSheet, list1);
                else if (data is List<Json2Order> list2) AddChartsJson2(chartSheet, list2);
                else if (data is List<Json3Order> list3) AddChartsJson3(chartSheet, list3);

                package.Save();
            }
        }


        // Chart Logik for JSON 1 dayly sales
        private void AddChartsJson1(ExcelWorksheet sheet, List<Json1Order> orders)
        {
            // Aggregierte Data per day
            var stats = orders.GroupBy(x => x.Created.Date)
                              .OrderBy(x => x.Key)
                              .Select(g => new { Date = g.Key, Total = g.Sum(x => x.Amount) })
                              .ToList();

            // Write helper data (Column Z)
            int startCol = 26;
            sheet.Cells[1, startCol].Value = "Datum";
            sheet.Cells[1, startCol + 1].Value = "Umsatz";
            sheet.Column(startCol).Style.Numberformat.Format = "dd.MM.yyyy";

            for (int i = 0; i < stats.Count; i++)
            {
                sheet.Cells[i + 2, startCol].Value = stats[i].Date;
                sheet.Cells[i + 2, startCol + 1].Value = stats[i].Total;
            }

            // Make line chart
            var chart = sheet.Drawings.AddChart("SalesLineChart", eChartType.Line);
            chart.Title.Text = "Täglicher Umsatzverlauf";
            chart.SetPosition(1, 0, 1, 0);
            chart.SetSize(800, 400);

            var series = chart.Series.Add(
                sheet.Cells[2, startCol + 1, stats.Count + 1, startCol + 1],
                sheet.Cells[2, startCol, stats.Count + 1, startCol]
            );

            series.Header = "Tagsumsatz";

            chart.XAxis.Title.Text = "Datum";
            chart.XAxis.Format = "dd.MM.yyyy";
            chart.XAxis.Title.Font.Size = 10;

            chart.YAxis.Title.Text = "Umsatz in €";
            chart.YAxis.Title.Font.Size = 10;
            chart.YAxis.Format = "#,##0 €";
            chart.Legend.Remove();
        }

        // Chart Logik for JSON 2 top products
        private void AddChartsJson2(ExcelWorksheet sheet, List<Json2Order> orders)
        {
            // Aggregierte Data per SKU
            var skuStats = orders.Where(o => o.Items != null)
                                 .SelectMany(o => o.Items)
                                 .GroupBy(i => i.Sku)
                                 .Select(g => new { Sku = g.Key, Total = g.Sum(i => i.ItemTotal) })
                                 .OrderByDescending(x => x.Total)
                                 .Take(10) // Top 10 products
                                 .ToList();

            int startCol = 26;
            for (int i = 0; i < skuStats.Count; i++)
            {
                sheet.Cells[i + 1, startCol].Value = skuStats[i].Sku;
                sheet.Cells[i + 1, startCol + 1].Value = skuStats[i].Total;
            }

            // Make bar chart
            var chart = sheet.Drawings.AddChart("ProductBarChart", eChartType.ColumnClustered);
            chart.Title.Text = "Top 10 Produkte (Umsatz)";
            chart.SetPosition(1, 0, 1, 0);
            chart.SetSize(600, 400);

            var series = chart.Series.Add(sheet.Cells[1, startCol + 1, skuStats.Count, startCol + 1], sheet.Cells[1, startCol, skuStats.Count, startCol]);
           
            series.Header = "Umsatzsumme";
            chart.XAxis.Title.Text = "Produkt (SKU)";
            chart.YAxis.Title.Text = "Umsatz in €";
            chart.YAxis.Format = "#,##0 €";
            chart.Legend.Remove();
        }

        // Chart Logik for JSON 3 sales rep and status
        private void AddChartsJson3(ExcelWorksheet sheet, List<Json3Order> orders)
        {
            // Chart 1: Sales Rep (Bar Chart)
            var repStats = orders.GroupBy(o => o.SalesRep)
                                 .Select(g => new { Name = string.IsNullOrEmpty(g.Key) ? "N/A" : g.Key, Total = g.Sum(x => x.TotalAmount) })
                                 .ToList();

            int startCol = 26;
            for (int i = 0; i < repStats.Count; i++)
            {
                sheet.Cells[i + 1, startCol].Value = repStats[i].Name;
                sheet.Cells[i + 1, startCol + 1].Value = repStats[i].Total;
            }

            // Make horizontal bar chart
            var barChart = sheet.Drawings.AddChart("RepChart", eChartType.BarClustered);
            barChart.Title.Text = "Umsatz pro Vertreter";
            barChart.SetPosition(1, 0, 1, 0);
            barChart.SetSize(500, 450);
            var series = barChart.Series.Add(sheet.Cells[1, startCol + 1, repStats.Count, startCol + 1], sheet.Cells[1, startCol, repStats.Count, startCol]);

            series.Header = "Gesamtumsatz";

            barChart.XAxis.Title.Text = "Vertreter";
            barChart.YAxis.Title.Text = "Umsatz in €";
            barChart.YAxis.Format = "#,##0 €";
            barChart.Legend.Remove();

            // Chart 2: Status (Pie Chart)
            var statusStats = orders.GroupBy(o => o.Status).Select(g => new { Status = g.Key, Count = g.Count() }).ToList();

            int statusCol = 29;
            for (int i = 0; i < statusStats.Count; i++)
            {
                sheet.Cells[i + 1, statusCol].Value = statusStats[i].Status;
                sheet.Cells[i + 1, statusCol + 1].Value = statusStats[i].Count;
            }

            var pieChart = sheet.Drawings.AddChart("StatusChart", eChartType.Doughnut);
            pieChart.Title.Text = "Status Übersicht";
            pieChart.SetPosition(1, 0, 10, 0);
            pieChart.SetSize(400, 400);
            var pieSeries = pieChart.Series.Add(sheet.Cells[1, statusCol + 1, statusStats.Count, statusCol + 1], sheet.Cells[1, statusCol, statusStats.Count, statusCol]);

            pieSeries.Header = "Anzahl Bestellungen";
            pieChart.ShowDataLabelsOverMaximum = true;
        }
    }
}
