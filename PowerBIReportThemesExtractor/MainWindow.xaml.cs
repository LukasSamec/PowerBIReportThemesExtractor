using PowerBIReportThemesExtractor.Extractor;
using System;
using System.Drawing;
using System.Windows;
using System.Windows.Forms;

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
                System.Windows.MessageBox.Show("Please, specify path to PowerBI file and path to theme", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            ThemeExtractor extractor = new ThemeExtractor(reportFileTextBox.Text, themeFileTextBox.Text, colorTextBox.Text);
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
            System.Windows.MessageBox.Show(ex.Message, "Error",MessageBoxButton.OK,MessageBoxImage.Error);
        }

        private void ChooseColor_Click(object sender, RoutedEventArgs e)
        {
            ColorDialog dialog = new ColorDialog();
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Color color = dialog.Color;
                String code = (color.ToArgb() & 0x00FFFFFF).ToString("X6");
                colorTextBox.Text = "#"+code;
                colorRectangle.Fill = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(color.R,color.G,color.B));
            }
        }
    }
}
