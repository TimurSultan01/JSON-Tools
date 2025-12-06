using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using JSON_Tools.Models;
using JSON_Tools.Services;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows;

namespace JSON_Tools.ViewModels
{
    public partial class MainViewModel : ObservableObject
    {
        private readonly JsonImportService _importService = new JsonImportService();
        private readonly ExcelExportService _exportService = new ExcelExportService();

        [ObservableProperty] private List<Json1Order>? _json1Orders;
        [ObservableProperty] private List<Json2Order>? _json2Orders;
        [ObservableProperty] private List<Json3Order>? _json3Orders;

        // Visibility controlls for different data grids
        [ObservableProperty] private Visibility _vis1 = Visibility.Collapsed;
        [ObservableProperty] private Visibility _vis2 = Visibility.Collapsed;
        [ObservableProperty] private Visibility _vis3 = Visibility.Collapsed;

        [ObservableProperty] private string _statusMessage = "Bitte Datei laden.";

        private object? _currentData;

        // Command to load orders from JSON
        [RelayCommand]
        private void LoadOrders(string filePath)
        {
            if (string.IsNullOrEmpty(filePath)) return;

            try
            {
                var rawData = _importService.LoadOrders(filePath);
                if (rawData == null) throw new InvalidOperationException("Keine Daten geladen.");

                // Perform validation and collect warnings
                string validationErrors = ValidateData(rawData);

                if (!string.IsNullOrEmpty(validationErrors))
                {
                    StatusMessage = "Daten mit Warnungen geladen.";
                    MessageBox.Show($"Folgende Probleme wurden gefunden:\n\n{validationErrors}",
                                    "Validierungshinweis",
                                    MessageBoxButton.OK,
                                    MessageBoxImage.Warning);
                }

                _currentData = rawData;

                Vis1 = Visibility.Collapsed;
                Vis2 = Visibility.Collapsed;
                Vis3 = Visibility.Collapsed;

                // Assign loaded data to the appropriate model property
                if (_currentData is List<Json1Order> list1)
                {
                    Json1Orders = list1.OrderBy(o => o.OrderId).ToList();
                    Vis1 = Visibility.Visible;
                    StatusMessage = $"Format 1 geladen ({list1.Count} Bestellungen)";
                }
                else if (_currentData is List<Json2Order> list2)
                {
                    Json2Orders = list2.OrderBy(o => o.OrderId).ToList();
                    Vis2 = Visibility.Visible;
                    StatusMessage = $"Format 2 geladen ({list2.Count} Bestellungen)";
                }
                else if (_currentData is List<Json3Order> list3)
                {
                    Json3Orders = list3.OrderBy(o => o.OrderId).ToList();
                    Vis3 = Visibility.Visible;
                    StatusMessage = $"Format 3 geladen ({list3.Count} Bestellungen)";
                }

                ExportOrdersCommand.NotifyCanExecuteChanged();
            }

            // Error handling for loading process
            catch (InvalidOperationException ex)
            {
                StatusMessage = "Formatfehler.";
                MessageBox.Show(ex.Message, "Format-Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (FileNotFoundException)
            {
                StatusMessage = "Datei nicht gefunden.";
                MessageBox.Show("Die angegebene Datei existiert nicht.", "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (JsonReaderException ex)
            {
                if (ex.Message.Contains("Could not convert") || ex.Message.Contains("Error reading") || ex.Message.Contains("Input string was not in a correct format"))
                {
                    StatusMessage = "Datentyp-Fehler.";

                    string feldInfo = string.IsNullOrEmpty(ex.Path) ? "Ein unbekanntes Feld" : $"Das Feld '{ex.Path}'";
                    string message = $"{feldInfo} enthält einen ungültigen Wert.\n" +
                                     $"Der Wert konnte nicht in den erwarteten Datentyp (z.B. Zahl oder Datum) umgewandelt werden.\n\n" +
                                     $"Ort: Zeile {ex.LineNumber}, Position {ex.LinePosition}";
                    MessageBox.Show(message, "Format-Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                else
                {
                    StatusMessage = "Ungültiges JSON.";
                    MessageBox.Show($"Die Datei enthält Syntaxfehler in Zeile {ex.LineNumber}.\n" +
                                    "Bitte überprüfen Sie auf fehlende Klammern oder Kommas.", 
                                    "Syntax-Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (NotSupportedException)
            {
                MessageBox.Show("Das Format der Datei wurde nicht erkannt.\n\n" +
                          "Mögliche Ursachen:\n" +
                          "- Felder fehlen (z.B. 'OrderId' oder 'price')\n" +
                          "- Unbekannte Felder sind enthalten\n" +
                          "- Datentypen sind falsch (z.B. Text statt Zahl)", "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                StatusMessage = "Kritischer Fehler.";
                MessageBox.Show($"Unerwarteter Fehler:\n{ex.Message}", "Systemfehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        // Warning validation for loaded data
        private string ValidateData(object data)
        {
            StringBuilder sb = new StringBuilder();
            HashSet<string> knownIds = new HashSet<string>();


            // Warning for Json1 Model
            if (data is List<Json1Order> json1Orders)
            {
                for (int i = 0; i < json1Orders.Count; i++)
                {
                    var item = json1Orders[i];

                    if (item.OrderId <= 0) sb.AppendLine($"Zeile {i + 1}: Ungültige OrderId.");
                    else if (!knownIds.Add(item.OrderId.ToString())) sb.AppendLine($"Zeile {i + 1}: Doppelte OrderId {item.OrderId}.");

                    if (string.IsNullOrWhiteSpace(item.Customer)) sb.AppendLine($"Zeile {i + 1}: Kunde fehlt.");
                    
                    if (item.Created > System.DateTime.Now) sb.AppendLine($"Zeile {i + 1}: Erstelldatum liegt in der Zukunft.");
                    if (item.Amount < 0) sb.AppendLine($"Zeile {i + 1}: Gesamtbetrag darf nicht negativ sein.");
                }
            }

            // Warning for Json2 Model
            else if (data is List<Json2Order> json2Orders)
            {
                for (int i = 0; i < json2Orders.Count; i++)
                {
                    var item = json2Orders[i];

                    if (item.OrderId <= 0) sb.AppendLine($"Zeile {i + 1}: Ungültige OrderId.");
                    else if (!knownIds.Add(item.OrderId.ToString())) sb.AppendLine($"Zeile {i + 1}: Doppelte OrderId {item.OrderId}.");

                    if (string.IsNullOrWhiteSpace(item.Customer)) sb.AppendLine($"Zeile {i + 1}: Kunde fehlt.");

                    if (item.Created > System.DateTime.Now) sb.AppendLine($"Zeile {i + 1}: Erstelldatum liegt in der Zukunft.");

                    if (item.Items == null || item.Items.Count == 0) sb.AppendLine($"Zeile {i + 1}: Es muss mindestens ein Item vorhanden sein.");
                    else
                    {
                        for (int j = 0; j < item.Items.Count; j++)
                        {
                            var subItem = item.Items[j];
                            if (string.IsNullOrWhiteSpace(subItem.Sku)) sb.AppendLine($"Zeile {i + 1}, Item {j + 1}: Sku fehlt.");
                            if (subItem.Qty <= 0) sb.AppendLine($"Zeile {i + 1}, Item {j + 1}: Menge muss größer als 0 sein.");
                            if (subItem.Price < 0) sb.AppendLine($"Zeile {i + 1}, Item {j + 1}: Preis darf nicht negativ sein.");
                        }
                    }
                }
            }

            // Warning for Json3 Model
            else if (data is List<Json3Order> json3Orders)
            {
                var validStatus = new HashSet<string> {"shipped", "processing", "cancelled" };

                for (int i = 0; i < json3Orders.Count; i++)
                {
                    var item = json3Orders[i];

                    if (item.OrderId <= 0) sb.AppendLine($"Zeile {i + 1}: Ungültige OrderId.");
                    else if (!knownIds.Add(item.OrderId.ToString())) sb.AppendLine($"Zeile {i + 1}: Doppelte OrderId {item.OrderId}.");

                    if (string.IsNullOrWhiteSpace(item.Customer)) sb.AppendLine($"Zeile {i + 1}: Kunde fehlt.");
                    if (string.IsNullOrWhiteSpace(item.SalesRep)) sb.AppendLine($"Zeile {i + 1}: Vertreter fehlt.");

                    if (item.Created > System.DateTime.Now) sb.AppendLine($"Zeile {i + 1}: Erstelldatum liegt in der Zukunft.");

                    if (string.IsNullOrWhiteSpace(item.Status) || !validStatus.Contains(item.Status.ToLower()))
                        sb.AppendLine($"Zeile {i + 1}: Ungültiger Status '{item.Status}'. Erlaubt: cancelled, shipped, processing");

                    if (item.Delivery == null) sb.AppendLine($"Zeile {i + 1}: Lieferinformationen fehlen.");
                    else
                    {
                        bool cancelled = item.Status.ToLower() == "cancelled";
                        bool hasDate = item.Delivery.DeliveryDate != null;
                        if (!hasDate && !cancelled) sb.AppendLine($"Zeile {i + 1}: Lieferdatum fehlt.");
                        else if (hasDate && !cancelled && item.Delivery.DeliveryDate < item.Created) 
                            sb.AppendLine($"Zeile {i + 1}: Lieferdatum liegt vor dem Erstelldatum.");
                       
                        if (string.IsNullOrWhiteSpace(item.Delivery.Address)) sb.AppendLine($"Zeile {i + 1}: Lieferadresse fehlt.");
                    }


                    if (item.Items == null || item.Items.Count == 0) sb.AppendLine($"Zeile {i + 1}: Es muss mindestens ein Item vorhanden sein.");
                    else
                    {
                        for (int j = 0; j < item.Items.Count; j++)
                        {
                            var subItem = item.Items[j];
                            if (string.IsNullOrWhiteSpace(subItem.Sku)) sb.AppendLine($"Zeile {i + 1}, Item {j + 1}: Sku fehlt.");
                            if (subItem.Qty <= 0) sb.AppendLine($"Zeile {i + 1}, Item {j + 1}: Menge muss größer als 0 sein.");
                            if (subItem.Price < 0) sb.AppendLine($"Zeile {i + 1}, Item {j + 1}: Preis darf nicht negativ sein.");
                        }
                    }
                }
            }

            return sb.ToString();
        }


        // Command to export orders to Excel
        [RelayCommand(CanExecute = nameof(CanExport))]
        private void ExportOrders(string filePath)
        {
            try
            {
                if (_currentData != null)
                {
                    _exportService.ExportToExcel(_currentData, filePath);
                    StatusMessage = "Export erfolgreich!";
                }
                else MessageBox.Show("Keine Daten zum Exportieren vorhanden.");
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Export Fehler: {ex.Message}");
            }
        }

        private bool CanExport() => _currentData != null;
    }
}