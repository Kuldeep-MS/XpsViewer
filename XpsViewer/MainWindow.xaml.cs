using System.Collections.ObjectModel;
using System.IO;
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
using System.Windows.Xps.Packaging;

namespace XpsViewer
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public ObservableCollection<string> Files = new ObservableCollection<string>();
        public MainWindow()
        {
            InitializeComponent();

            Directory.EnumerateFiles(@"./xps", "*.xps").Select(System.IO.Path.GetFullPath).ToList().ForEach(Files.Add);
            cmbBox.ItemsSource = Files;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (cmbBox.SelectedItem == null) return;

            var selectedItem = cmbBox.SelectedItem.ToString();
            //if (string.IsNullOrEmpty(txtBox.Text)) return;
            try
            {

                System.Diagnostics.Debug.WriteLine($"[DEBUG]Loading.. {selectedItem}");
                System.Diagnostics.Trace.WriteLine($"Loading.. {selectedItem}");
                XpsDocument myDoc = new XpsDocument(selectedItem, FileAccess.Read);

                var a = myDoc.GetFixedDocumentSequence();
                dv.Document = a;

                txtBlock.Text = $"{selectedItem} Loaded successfully";
            }
            catch (Exception ex)
            {
                txtBlock.Text = $"{selectedItem} Failed ({ex.Message})";

            }

        }
    }
}