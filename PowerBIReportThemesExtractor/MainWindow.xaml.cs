using Microsoft.Win32;
using PowerBIReportThemesExtractor.Extractor;
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

namespace PowerBIReportThemesExtractor
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

        private void LoadFIleButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.ShowDialog();
            //dialog.Filter = "Power BI file (*.pbix)";
            reportFIleTextBox.Text = dialog.FileName;
            extractButton.IsEnabled = true;           
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            ThemeExtractor extractor = new ThemeExtractor(reportFIleTextBox.Text);
        }
    }
}
