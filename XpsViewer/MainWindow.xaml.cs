extern alias WinDep;
using PackageDependency = WinDep::Microsoft.Windows.ApplicationModel.DynamicDependency.PackageDependency;

using Microsoft.Windows.ApplicationModel.DynamicDependency;
using Microsoft.Windows.SemanticSearch;


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
using System;
using System.Reflection.Metadata;
using System.Reflection;
using System.Xml.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Drawing;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Shapes;
using System.Windows.Shapes;
using System.Windows.Controls;

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
                StringBuilder documentText = new StringBuilder();
                FixedDocumentSequence fds = dv.Document as FixedDocumentSequence;
            
                if (fds != null)
                {
                    foreach (DocumentReference dr in fds.References)
                    {
                        FixedDocument fd = dr.GetDocument(false);
                        foreach (PageContent pc in fd.Pages)
                        {
                            FixedPage fp = pc.GetPageRoot(false);
                            foreach (UIElement element in fp.Children)
                            {
                                if (element is Glyphs glyphs)
                                {
                                    documentText.AppendLine(glyphs.UnicodeString);
                                }
                            }
                        }
                    }
                }

                var docText = documentText.ToString();
                
                var minVersion = new Windows.ApplicationModel.PackageVersion(0, 2, 0, 0);
                var dep = PackageDependency.Create("WindowsWorkload.Manager.1_8wekyb3d8bbwe", minVersion);
                dep.Add();

                dep = PackageDependency.Create("WindowsWorkload.Client.1_8wekyb3d8bbwe", minVersion);
                dep.Add();
                // add winappsdk
                            minVersion = new global::Windows.ApplicationModel.PackageVersion(
                                Microsoft.WindowsAppSDK.Runtime.Version.Major,
                                Microsoft.WindowsAppSDK.Runtime.Version.Minor,
                                Microsoft.WindowsAppSDK.Runtime.Version.Build,
                                Microsoft.WindowsAppSDK.Runtime.Version.Revision);

                dep = PackageDependency.Create(Microsoft.WindowsAppSDK.Runtime.Packages.Framework.PackageFamilyName, minVersion);
                dep.Add();

                dv.SetDocText(docText);
                this.MouseLeftButtonDown += DocumentViewer_MouseLeftButtonDown;

                //  var bitmap = dv.GetBitmapSource();
                //var paths = dv.GetPaths();
                // showImages(bitmap, paths, dv.canvases);

            }
            catch (Exception ex)
            {
                txtBlock.Text = $"{selectedItem} Failed ({ex.Message})";
            }

             

        }

        private void DocumentViewer_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            dv.MouseLeftButtonDown2();
/*            if (highlightRect != null)
            {
                canvases[rankImages[currentIndexImages]].Children.Remove(highlightRect);
                highlightRect = null;
            }
*/
        }

        private void SearchToggle_Checked(object sender, RoutedEventArgs e)
        {
            dv.IsTextSearch = false;
        }

        private void SearchToggle_Unchecked(object sender, RoutedEventArgs e)
        {
            dv.IsTextSearch = true;
        }


        private async Task showImages(List<BitmapSource> images, List<System.Windows.Shapes.Path> paths, List<System.Windows.Controls.Canvas> canvases)
        {
            int i = 0;
            foreach (BitmapSource bit in images)
            {
              //  myImage.Source = bit;
                // Add wait for 3 seconds below

                var region = paths[i].Data.Bounds;
                // Create a semi-transparent rectangle to represent the highlight.
                System.Windows.Shapes.Rectangle highlight = new System.Windows.Shapes.Rectangle
                {
                    Width = region.Width,
                    Height = region.Height,
                    Fill = new SolidColorBrush(System.Windows.Media.Color.FromArgb(128, 0, 0, 255)) // Semi-transparent blue
                };


                canvases[i].Children.Add(highlight);
                    System.Windows.Controls.Canvas.SetLeft(highlight, System.Windows.Controls.Canvas.GetLeft(paths[i]));
                    System.Windows.Controls.Canvas.SetTop(highlight, System.Windows.Controls.Canvas.GetTop(paths[i]));

                // Position the highlight rectangle at the specified coordinates.
                System.Windows.Controls.Canvas.SetLeft(highlight, region.Left);
                System.Windows.Controls.Canvas.SetTop(highlight, region.Top);
                i++;

                /*// Get the private field _contentHost from the DocumentViewer class
                var contentHostField = typeof(DocumentViewer).GetField("_contentHost", BindingFlags.NonPublic | BindingFlags.Instance);

                // Get the value of the _contentHost field for the documentViewer instance
                var contentHost = contentHostField.GetValue(dv);

                // Check if the contentHost is a Canvas
                if (contentHost is Canvas canvas)
                {
                    // Add the child element to the Canvas
                    canvas.Children.Add(highlight);
                }*/

                // Add the highlight rectangle to the DocumentViewer.
                // This assumes that the DocumentViewer is contained within a Canvas.
                //((Canvas)Parent).Children.Add(highlight);
                await Task.Delay(3000);
                canvases[i-1].Children.Remove(highlight);
                
            }
        }
    }
}