using Microsoft.Win32;
using PowerBIReportThemesExtractor.Extractor;
using System;
using System.Windows;


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
            AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);     
        }
    
        private void LoadReportFileButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Power BI file (*.pbix) | *.pbix|Power BI template file (*.pbit)|*.pbit";
            dialog.ShowDialog();                  
            reportFileTextBox.Text = dialog.FileName;           
        }

        private void ExtractButton_Click(object sender, RoutedEventArgs e)
        {
            if (String.IsNullOrEmpty(reportFileTextBox.Text) && String.IsNullOrEmpty(themeFileTextBox.Text))
            {
                MessageBox.Show("Please, specify path to PowerBI file and path to theme", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            ThemeExtractor extractor = new ThemeExtractor(reportFileTextBox.Text, themeFileTextBox.Text);
            extractor.Extract();            
        }

        private void SaveThemeFileButton_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.Filter = "Json files (*.json)|*.json";
            dialog.ShowDialog();
            themeFileTextBox.Text = dialog.FileName;
        }

        static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            Exception ex = (Exception)e.ExceptionObject;
            MessageBox.Show(ex.Message, "Error",MessageBoxButton.OK,MessageBoxImage.Error);
        }
    }
}
