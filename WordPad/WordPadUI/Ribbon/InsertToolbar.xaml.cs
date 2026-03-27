using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text;
using System.Threading.Tasks;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.Storage.Streams;
using Windows.Storage;
using Windows.System;
using Windows.UI.Text;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Media.Imaging;
using Windows.UI.Xaml.Navigation;
using Windows.Storage.Pickers;
using Windows.ApplicationModel.Core;
using Windows.Management.Deployment;

// Szablon elementu Kontrolka użytkownika jest udokumentowany na stronie https://go.microsoft.com/fwlink/?LinkId=234236

namespace WordPad.WordPadUI.Ribbon
{
    public sealed partial class InsertToolbar : UserControl
    {
        public RichEditBox Editor { get; set; }
        public InsertToolbar()
        {
            this.InitializeComponent();
        }

        private async void InsertDateAndTimeButton_Click(object sender, RoutedEventArgs e)
        {
            ContentDialog dialog = new ContentDialog();
            dialog.Title = "Insert current date and time";

            // Create a ListView for the user to select the date format
            ListView listView = new ListView();
            listView.SelectionMode = ListViewSelectionMode.Single;

            // Create a list of date formats to display in the ListView
            List<string> dateFormats = new List<string>();
            dateFormats.Add(DateTime.Now.ToString("dd.MM.yyyy"));
            dateFormats.Add(DateTime.Now.ToString("dd MMM yyyy"));
            dateFormats.Add(DateTime.Now.ToString("dddd, dd MMMM yyyy"));
            dateFormats.Add(DateTime.Now.ToString("dd MMMM yyyy"));
            dateFormats.Add(DateTime.Now.ToString("HH:mm:ss"));

            // Set the ItemsSource of the ListView to the list of date formats
            listView.ItemsSource = dateFormats;

            // Set the content of the ContentDialog to the ListView
            dialog.Content = listView;

            // Make the insert button colored
            dialog.DefaultButton = ContentDialogButton.Primary;

            // Add an "Insert" button to the ContentDialog
            dialog.PrimaryButtonText = "OK";
            dialog.PrimaryButtonClick += (s, args) =>
            {
                string selectedFormat = listView.SelectedItem as string;
                try
                {
                    string formattedDate = dateFormats[listView.SelectedIndex];
                    Editor.Document.Selection.Text = formattedDate;
                    Editor.Focus(FocusState.Keyboard);
                }
                catch
                {

                }
            };

            // Add a "Cancel" button to the ContentDialog
            dialog.SecondaryButtonText = "Cancel";

            // Show the ContentDialog
            await dialog.ShowAsync();
        }

        static async Task<AppListEntry> GetAppByPackageFamilyNameAsync(string packageFamilyName)
        {
            var pkgManager = new PackageManager();
            var pkg = pkgManager.FindPackagesForUser("", packageFamilyName).FirstOrDefault();

            if (pkg == null) return null;

            var apps = await pkg.GetAppListEntriesAsync();
            var firstApp = apps.FirstOrDefault();
            return firstApp;
        }

        private async void InsertObjectButton_Click(object sender, RoutedEventArgs e)
        {
            // Create a ContentDialog
            ContentDialog dialog = new ContentDialog();
            dialog.Title = "Insert object";

            // Create a ListView for the user to select the insert option
            ListView listView = new ListView();
            listView.SelectionMode = ListViewSelectionMode.Single;

            // Create a list of insert options to display in the ListView
            List<string> insertOptions = new List<string>();
            //insertOptions.Add("Paint a picture");
            insertOptions.Add("Table");

            // Set the ItemsSource of the ListView to the list of insert options
            listView.ItemsSource = insertOptions;

            // Set the content of the ContentDialog to the ListView
            dialog.Content = listView;

            // Make the default button BLU
            dialog.DefaultButton = ContentDialogButton.Primary;

            // Add an "Insert" button to the ContentDialog
            dialog.PrimaryButtonText = "OK";
            dialog.PrimaryButtonClick += async (s, args) =>
            {
                string selectedOption = listView.SelectedItem as string;

                // Draw an image using mspaint and insert it into the RichEditBox
                if (selectedOption == "Paint a picture")
                {
                    var app = await GetAppByPackageFamilyNameAsync("Microsoft.Paint_8wekyb3d8bbwe");

                    if (app != null)
                    {
                        await app.LaunchAsync();
                    }

                    // Wait for the user to draw an image and save it to a temporary file
                    StorageFile tempFile = await ApplicationData.Current.TemporaryFolder.CreateFileAsync("temp.bmp", CreationCollisionOption.ReplaceExisting);
                    while (true)
                    {
                        await Task.Delay(TimeSpan.FromSeconds(6));
                        if (await tempFile.GetBasicPropertiesAsync() != null) break;
                    }

                    // Load the image into a BitmapImage object
                    BitmapImage bitmapImage = new BitmapImage();
                    using (IRandomAccessStream stream = await tempFile.OpenAsync(FileAccessMode.Read))
                    {
                        bitmapImage.SetSource(stream);
                    }

                    // Insert the image into the RichEditBox
                    using (IRandomAccessStream stream = await tempFile.OpenAsync(FileAccessMode.Read))
                    {
                        IRandomAccessStreamReference imageStreamRef = RandomAccessStreamReference.CreateFromStream(stream);
                        using (var imageStream = await imageStreamRef.OpenReadAsync())
                        {
                            // Insert the object into the RichEditBox
                            Editor.Document.Selection.InsertImage(200, 200, 0, VerticalCharacterAlignment.Baseline, "img", imageStream);

                        }
                    }

                    // Delete the temporary file
                    await tempFile.DeleteAsync();
                }
                // Insert a table into the RichEditBox
                else if (selectedOption == "Table")
                {
                    //CreateStringBuilder object
                    StringBuilder strTable = new StringBuilder();

                    //Beginning of rich text format,don’t alter this line
                    strTable.Append(@"{\rtf1 ");

                    //Create 5 rows with 4 columns
                    for (int i = 0; i < 5; i++)
                    {
                        //Start the row
                        strTable.Append(@"\trowd");

                        //First cell with width 1000.
                        strTable.Append(@"\cellx1000");

                        //Second cell with width 1000.Ending point is 2000, which is 1000+1000.
                        strTable.Append(@"\cellx2000");

                        //Third cell with width 1000.Endingat3000,which is 2000+1000.
                        strTable.Append(@"\cellx3000");

                        //Last cell with width 1000.Ending at 4000 (which is 3000+1000)
                        strTable.Append(@"\cellx4000");

                        //Append the row in StringBuilder
                        strTable.Append(@"\intbl \cell \row"); //create the row
                    }

                    strTable.Append(@"\pard");

                    strTable.Append(@"}");

                    var strTableString = strTable.ToString();


                    Editor.Document.Selection.SetText(TextSetOptions.FormatRtf, strTableString);
                }
            };

            // Add a "Cancel" button to the ContentDialog
            dialog.SecondaryButtonText = "Cancel";

            // Show the ContentDialog
            await dialog.ShowAsync();
        }

        private async void InsertImageButton_Click(object sender, RoutedEventArgs e)
        {
            // Open an image file.
            FileOpenPicker open = new FileOpenPicker();
            open.SuggestedStartLocation = PickerLocationId.DocumentsLibrary;
            open.FileTypeFilter.Add(".png");
            open.FileTypeFilter.Add(".jpg");
            open.FileTypeFilter.Add(".jpeg");

            StorageFile file = await open.PickSingleFileAsync();

            if (file != null)
            {
                IRandomAccessStream randAccStream = await file.OpenAsync(FileAccessMode.Read);
                var properties = await file.Properties.GetImagePropertiesAsync();
                int width = (int)properties.Width;
                int height = (int)properties.Height;

                // Load the file into the Document property of the RichEditBox.
                Editor.Document.Selection.InsertImage(width, height, 0, VerticalCharacterAlignment.Baseline, "img", randAccStream);
                Editor.Focus(FocusState.Keyboard);
            }
        }

        private async void InsertPaintImageButton_Click(object sender, RoutedEventArgs e)
        {
            var app = await GetAppByPackageFamilyNameAsync("Microsoft.Paint_8wekyb3d8bbwe");

            if (app != null)
            {
                await app.LaunchAsync();
            }
        }
    }
}
