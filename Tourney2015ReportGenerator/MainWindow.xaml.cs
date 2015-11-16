using FileHelpers;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Tourney2015ReportGenerator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void InputFileTextBox_Changed(object sender, TextChangedEventArgs e)
        {
            generateReportButton.IsEnabled = false;

            var filename = inputFileTextBox.Text;
            if (File.Exists(filename))
            {
                generateReportButton.IsEnabled = true;
            }
        }

        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog()
            {
                CheckFileExists = true,
                CheckPathExists = true,
                Filter = "CSV File (*.csv)|*.csv"
            };

            var result = dialog.ShowDialog();
            if (result.GetValueOrDefault())
            {
                inputFileTextBox.Text = dialog.FileName;
            }
        }

        private async void GenerateReportButton_Click(object sender, RoutedEventArgs e)
        {
            progressBar.Visibility = Visibility.Visible;
            this.IsEnabled = false;
            var inputFileName = inputFileTextBox.Text;
            await Task.Run(() =>
            {
                try
                {
                    var records = ReadRegistrationFile(inputFileName);
                    var reporter = new RegistrationReporter(records);
                    var reportExporter = new ExcelReportPrinter(reporter);

                    var saveDialog = new SaveFileDialog()
                    {
                        OverwritePrompt = true,
                        Filter = "Excel File (*.xlsx)|*.xlsx"
                    };
                    var result = saveDialog.ShowDialog();
                    if (result.GetValueOrDefault())
                    {
                        var outputFileName = saveDialog.FileName;
                        try
                        {
                            reportExporter.CreateReport(outputFileName);
                            MessageBox.Show("Report created successfully");
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(string.Format("Unable to save file: {0}", ex.Message));
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(string.Format("Unable to read file: {0}", ex.Message));
                }
            });

            progressBar.Visibility = Visibility.Hidden;
            this.IsEnabled = true;
        }

        private IEnumerable<RegistrationRecord> ReadRegistrationFile(string inputFile)
        {
            var engine = new FileHelperEngine<RegistrationRecord>();
            return engine.ReadFile(inputFile);
        }
    }
}
