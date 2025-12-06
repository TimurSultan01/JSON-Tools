using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System.Collections.Generic;
using System.Windows;
using JSON_Tools.Models;
using JSON_Tools.Services;

namespace JSON_Tools.ViewModels
{
    public partial class MainViewModel : ObservableObject
    {
        private readonly JsonImportService _importService = new JsonImportService();
        private readonly ExcelExportService _exportService = new ExcelExportService();

        [ObservableProperty] private List<Json1Order> _json1Orders;
        [ObservableProperty] private List<Json2Order> _json2Orders;
        [ObservableProperty] private List<Json3Order> _json3Orders;

        // Sichtbarkeits-Steuerung
        [ObservableProperty] private Visibility _vis1 = Visibility.Collapsed;
        [ObservableProperty] private Visibility _vis2 = Visibility.Collapsed;
        [ObservableProperty] private Visibility _vis3 = Visibility.Collapsed;

        [ObservableProperty] private string _statusMessage = "Bitte Datei laden.";

        private object _currentData;

        // Command to load orders from JSON
        [RelayCommand]
        private void LoadOrders(string filePath)
        {
            if (string.IsNullOrEmpty(filePath)) return;

            try
            {
                _currentData = _importService.LoadOrders(filePath);

                Vis1 = Visibility.Collapsed;
                Vis2 = Visibility.Collapsed;
                Vis3 = Visibility.Collapsed;

                if (_currentData is List<Json1Order> list1)
                {
                    Json1Orders = list1;
                    Vis1 = Visibility.Visible;
                    StatusMessage = $"Format 1 geladen ({list1.Count} Bestellungen)";
                }
                else if (_currentData is List<Json2Order> list2)
                {
                    Json2Orders = list2;
                    Vis2 = Visibility.Visible;
                    StatusMessage = $"Format 2 geladen ({list2.Count} Bestellungen)";
                }
                else if (_currentData is List<Json3Order> list3)
                {
                    Json3Orders = list3;
                    Vis3 = Visibility.Visible;
                    StatusMessage = $"Format 3 geladen ({list3.Count} Bestellungen)";
                }

                ExportOrdersCommand.NotifyCanExecuteChanged();
            }
            catch (System.Exception ex)
            {
                StatusMessage = "Fehler beim Laden.";
                MessageBox.Show(ex.Message);
            }
        }

        // Command to export orders to Excel
        [RelayCommand(CanExecute = nameof(CanExport))]
        private void ExportOrders(string filePath)
        {
            try
            {
                _exportService.ExportToExcel(_currentData, filePath);
                StatusMessage = "Export erfolgreich!";
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private bool CanExport() => _currentData != null;
    }
}