using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System.Collections.ObjectModel;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using JSON_Tools.Models;
using JSON_Tools.Services;

namespace JSON_Tools.ViewModels
{
    public partial class MainViewModel : ObservableObject
    {
        private readonly JsonImportService _importService = new JsonImportService();
        private readonly ExcelExportService _exportService = new ExcelExportService();

        [ObservableProperty]
        private ObservableCollection<Order> _orders = new ObservableCollection<Order>();

        [ObservableProperty]
        private string _statusMessage = "Bereit. Bitte Datei laden.";

        // Command to load orders from JSON
        [RelayCommand]
        private void LoadOrders(string filePath)
        {
            if (string.IsNullOrEmpty(filePath)) return;

            try
            {
                var data = _importService.LoadOrders(filePath);

                Orders.Clear();
                foreach (var item in data) Orders.Add(item);

                StatusMessage = $"{Orders.Count} Einträge geladen.";
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
                _exportService.ExportToExcel(new List<Order>(Orders), filePath);
                StatusMessage = "Export erfolgreich!";
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private bool CanExport() => Orders.Count > 0;
    }
}