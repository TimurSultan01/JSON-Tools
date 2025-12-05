using JSON_Tools.ViewModels;
using Microsoft.Win32;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace JSON_Tools
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly MainViewModel _viewModel;

        public MainWindow()
        {
            InitializeComponent();
            _viewModel = new MainViewModel();
            this.DataContext = _viewModel;
        }

        private void LoadButton_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*",
                Title = "Wählen Sie die JSON-Datei zum Laden aus"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                _viewModel.LoadOrdersCommand.Execute(openFileDialog.FileName);
            }
        }

        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            if (!_viewModel.ExportOrdersCommand.CanExecute(null)) return;

            var saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                Title = "Speichern unter",
                FileName = "export_orders.xlsx"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                _viewModel.ExportOrdersCommand.Execute(saveFileDialog.FileName);
            }
        }
    }
}