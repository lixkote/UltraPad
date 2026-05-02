using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Graphics.Canvas.Text;
using Microsoft.Toolkit.Uwp.Helpers;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;
using Windows.ApplicationModel.Core;
using Windows.ApplicationModel.DataTransfer;
using Windows.ApplicationModel.Email;
using Windows.ApplicationModel.Resources.Core;
using Windows.Foundation;
using Windows.Graphics.Display;
using Windows.Graphics.Printing;
using Windows.Management.Deployment;
using Windows.Storage;
using Windows.Storage.AccessCache;
using Windows.Storage.Pickers;
using Windows.Storage.Provider;
using Windows.Storage.Streams;
using Windows.System;
using Windows.UI;
using Windows.UI.Core;
using Windows.UI.Core.Preview;
using Windows.UI.Text;
using Windows.UI.ViewManagement;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Documents;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Media.Animation;
using Windows.UI.Xaml.Media.Imaging;
using Windows.UI.Xaml.Navigation;
using WordPad.Helpers;
using WordPad.WordPadUI;
using WordPad.WordPadUI.Settings;
using static System.Net.Mime.MediaTypeNames;
using Application = Windows.UI.Xaml.Application;
using CheckBox = Windows.UI.Xaml.Controls.CheckBox;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using UnicodeEncoding = Windows.Storage.Streams.UnicodeEncoding;

// UltraPad made by Lixkote 
// Main page c# source code

namespace RectifyPad
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {
        private bool _isDirty = false;
        private string _lastSavedText = "";
        private string _currentFilePath = null;

        private bool _initialized = false;
        public StorageFile RichEditFile;
        private string appTitleStr => "UltraPad";

        ApplicationDataContainer localSettings = ApplicationData.Current.LocalSettings;

        private bool updateFontFormat = true;
        public string ZoomString => ZoomSlider.Value.ToString() + "%";

        private string fileNameWithPath = "";

        private double baseRulerWidth;

        string originalDocText = "";

        UnitManager unitConverter = new UnitManager();
        SettingsManagerMain settingsManager = new SettingsManagerMain();
        OdtHelper odtHelper = new OdtHelper();

        public List<string> Fonts
        {
            get
            {
                return CanvasTextFormat.GetSystemFontFamilies().OrderBy(f => f).ToList();
            }
        }

        public ObservableCollection<double> ZoomOptions { get; } = new ObservableCollection<double> { 5, 4, 3, 2, 1, 0.75, 0.5, 0.25, 0.125 };

        public List<double> FontSizes { get; } = new List<double>()
        {
                8,
                9,
                10,
                11,
                12,
                14,
                16,
                18,
                20,
                24,
                28,
                36,
                48,
                72
        };

        private void SetLineSpacing(double lineSpacingValue, LineSpacingRule lineSpacingRule = LineSpacingRule.Exactly)
        {
            // Get the document from the RichEditBox
            var document = Editor.Document;

            // Select the entire document (or specify a different range if needed)
            document.Selection.Expand(TextRangeUnit.Paragraph);

            // Modify the line spacing using ITextParagraphFormat
            var paragraphFormat = document.Selection.ParagraphFormat;

            // Use SetLineSpacing to set both the rule and the spacing value
            paragraphFormat.SetLineSpacing(lineSpacingRule, (float)lineSpacingValue);
        }

        public MainPage()
        {
            /////
            /// Startup Procedure
            /////

            // Enable navigation cache
            this.NavigationCacheMode = Windows.UI.Xaml.Navigation.NavigationCacheMode.Enabled;

            // Run the startup functions
            InitializeComponent();
            settingsManager.InitializeDefaults();
            LoadThemeFromSettings();
            LoadSettingsValues();
            PopulateRecents();
            ConnectRibbonToolbars();

            // ParagraphMenuIcon.FontFamily = (Windows.UI.Xaml.Media.FontFamily)Application.Current.Resources["CustomIconFont"];
            MenuParagraphIcon.FontFamily = (Windows.UI.Xaml.Media.FontFamily)Application.Current.Resources["CustomIconFont"];
            ParagraphIconHost.FontFamily = (Windows.UI.Xaml.Media.FontFamily)Application.Current.Resources["CustomIconFont"];

            SystemNavigationManagerPreview.GetForCurrentView().CloseRequested += OnCloseRequest;
            ribbongrid.DataContext = this;

            // Load the saved settings and apply them
            if (localSettings.Values["IsDarkThemeEditor"] != null)
            {
                EditorContainer.RequestedTheme = (bool)Windows.Storage.ApplicationData.Current.LocalSettings.Values["IsDarkThemeEditor"] ? ElementTheme.Dark : ElementTheme.Light;
            }
            if (localSettings.Values["isSpellCheckEnabled"] != null)
            {
                Editor.IsSpellCheckEnabled = (bool)Windows.Storage.ApplicationData.Current.LocalSettings.Values["isSpellCheckEnabled"] ? true : false;
            } 
            if (localSettings.Values["isTextPredictEnabled"] != null)
            {
                Editor.IsTextPredictionEnabled = (bool)Windows.Storage.ApplicationData.Current.LocalSettings.Values["isTextPredictEnabled"] ? true : false;
            }
            // Subscribe to theme change events
            SettingsPageManager.ThemeChanged += ChangeEditorContainerTheme;
            DataTransferManager dataTransferManager = DataTransferManager.GetForCurrentView();
            dataTransferManager.DataRequested += DataTransferManager_DataRequested;
            SetLineSpacing(1, LineSpacingRule.Multiple); // Single spacing


            ResetDirtyAfterDelay();
        }

        private void MarkClean()
        {
            _isDirty = false;
        }

        private void Editor_TextChanged(object sender, RoutedEventArgs e)
        {
            _isDirty = true;
        }
        private async void ResetDirtyAfterDelay()
        {
            await Task.Delay(500);
            _isDirty = false;
        }

        public void ChangeEditorContainerTheme(bool isDarkThemeEditor)
        {
            EditorContainer.RequestedTheme = isDarkThemeEditor ? ElementTheme.Dark : ElementTheme.Light;
            PrintSubItem.IsEnabled = isDarkThemeEditor ? false : true;
        }

        public void EnableEditorSpellCheck(bool isSpellCheckEnabled)
        {
            Editor.IsSpellCheckEnabled = isSpellCheckEnabled ? true : false;
        }

        public void EnableEditorAutocorrect(bool isTextPredictEnabled)
        {
            Editor.IsTextPredictionEnabled = isTextPredictEnabled ? true : false;
        }

        private void LoadSettingsValues()
        {
            try
            {
                // Load text wrapping value from settings:
                string textWrapping = localSettings.Values["textwrapping"] as string;
                if (textWrapping == "wrapwindow")
                {
                    Editor.TextWrapping = TextWrapping.Wrap;
                }
                else if (textWrapping == "nowrap")
                {
                    Editor.TextWrapping = TextWrapping.NoWrap;
                }
                else if (textWrapping == "wrapruler")
                {
                    // Add a function here that will do the ruler-based wrapping
                }

                // Load margin values from the settings:
                var settings = ApplicationData.Current.LocalSettings;

                string unit = settings.Values["unitSetting"] as string;
                string Lmargin = settings.Values["pagesetupLmargin"] as string;
                string Rmargin = settings.Values["pagesetupRmargin"] as string;
                string Tmargin = settings.Values["pagesetupTmargin"] as string;
                string Bmargin = settings.Values["pagesetupBmargin"] as string;

                // Debugging output to check retrieved values and their types
                Debug.WriteLine($"unit: {unit}, Lmargin: {Lmargin}, Rmargin: {Rmargin}, Tmargin: {Tmargin}, Bmargin: {Bmargin}");

                // Check if any of the values retrieved are null or not of type string
                if (!string.IsNullOrEmpty(unit) && !string.IsNullOrEmpty(Lmargin) && !string.IsNullOrEmpty(Rmargin) && !string.IsNullOrEmpty(Tmargin) && !string.IsNullOrEmpty(Bmargin))
                {
                    // Convert margin values to match the unit and format them as needed
                    // double left = unitConverter.ConvertToUnitAndFormat(Lmargin, unit);
                    // double right = unitConverter.ConvertToUnitAndFormat(Rmargin, unit);
                    // double top = unitConverter.ConvertToUnitAndFormat(Tmargin, unit);
                    // double bottom = unitConverter.ConvertToUnitAndFormat(Bmargin, unit);
                }
                else
                {
                    // Handle the case where one or more values are missing or not of type string
                    Debug.WriteLine("One or more settings values are missing or not of type string.");
                }
            }
            catch (Exception ex)
            {
                // Handle the exception
                Debug.WriteLine($"An exception occurred: {ex.Message}");
            }
        }


        private void LoadThemeFromSettings()
        {
            string value = (string)Windows.Storage.ApplicationData.Current.LocalSettings.Values["themeSetting"];
            if (value != null)
            {
                try
                {
                    // Change title bar color if needed
                    ApplicationViewTitleBar titleBar = ApplicationView.GetForCurrentView().TitleBar;
                    if (value == "Dark")
                    {
                        titleBar.ButtonForegroundColor = Colors.White;
                        App.RootTheme = ElementTheme.Dark;
                    }
                    else if (value == "Light")
                    {
                        titleBar.ButtonForegroundColor = Colors.Black;
                        App.RootTheme = ElementTheme.Light;
                    }
                    else
                    {
                        App.RootTheme = ElementTheme.Default;
                        if (Application.Current.RequestedTheme == ApplicationTheme.Dark)
                        {
                            titleBar.ButtonForegroundColor = Colors.White;
                        }
                        else
                        {
                            titleBar.ButtonForegroundColor = Colors.Black;
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Handle the exception
                    Debug.WriteLine($"An exception occurred: {ex.Message}");
                }

            }
            Window.Current.SetTitleBar(AppTitleBar);
        }

        private void ConnectRibbonToolbars()
        {
            editribbontoolbar.Editor = Editor;
            insertribbontoolbar.Editor = Editor;
            pararibbontoolbar.Editor = Editor;
            fontribbontoolbar.Editor = Editor;

            //collapsed variants also need to be 'connected'
            editribbontoolbarcol.Editor = Editor;
            insertribbontoolbarcol.Editor = Editor;
            pararibbontoolbarcol.Editor = Editor;
            pararibbontoolbarcol.MainPagea = this;
            fontribbontoolbarcol.Editor = Editor;
            TextRuler.editor = Editor;
        }

        private async void PopulateRecents()
        {
            var recentlyUsedItems = await RecentlyUsedHelper.GetRecentlyUsedItems();
            var recentItemsSubItem = RecentItemsSubItem;
            foreach (var item in recentlyUsedItems)
            {
                var menuItem = new MenuFlyoutItem { Text = item.Name };
                menuItem.Click += async (s, args) =>
                {
                    var file = await StorageFile.GetFileFromPathAsync(item.Path);
                    await RecentlyUsedHelper.AddToMostRecentlyUsedList(file);
                    // Open the file here
                };
                recentItemsSubItem.Items.Add(menuItem);
            }
        }

        private MarkerType _type = MarkerType.Bullet;


        private void MyListButton_IsCheckedChanged(Microsoft.UI.Xaml.Controls.ToggleSplitButton sender, Microsoft.UI.Xaml.Controls.ToggleSplitButtonIsCheckedChangedEventArgs args)
        {
            if (sender.IsChecked)
            {
                //add bulleted list
                Editor.Document.Selection.ParagraphFormat.ListType = _type;
            }
            else
            {
                //remove bulleted list
                Editor.Document.Selection.ParagraphFormat.ListType = MarkerType.None;
        
            }
        }

        string FixRtf(string rtf)
        {
            var output = new StringBuilder();

            int i = 0;

            while (i < rtf.Length)
            {
                int pictIndex = rtf.IndexOf(@"\pict", i);

                if (pictIndex == -1)
                {
                    output.Append(rtf.Substring(i));
                    break;
                }

                // append wszystko przed obrazem
                output.Append(rtf.Substring(i, pictIndex - i));

                // znajdź początek bloku {
                int start = rtf.LastIndexOf('{', pictIndex);

                if (start == -1)
                {
                    i = pictIndex + 5;
                    continue;
                }

                // znajdź koniec bloku licząc klamry
                int depth = 0;
                int end = start;

                for (; end < rtf.Length; end++)
                {
                    if (rtf[end] == '{') depth++;
                    else if (rtf[end] == '}') depth--;

                    if (depth == 0)
                    {
                        end++;
                        break;
                    }
                }

                string pictBlock = rtf.Substring(start, end - start);

                if (pictBlock.Contains(@"\wmetafile"))
                {
                    output.Append(@"{\pard\plain\fs20\b [Outdated WMF image format not supported in UltraPad]\par}");
                }
                else if (pictBlock.Contains(@"\emfblip"))
                {
                    output.Append(@"{\pard\plain\fs20\b [Outdated EMF image format not supported in UltraPad]\par}");
                }
                else
                {
                    output.Append(pictBlock);
                }

                i = end;
            }

            return output.ToString();
        }
        string EscapeRtf(string text)
        {
            return text
                .Replace(@"\", @"\\")
                .Replace("{", @"\{")
                .Replace("}", @"\}");
        }

        string ParseImage(XmlNode node, ZipArchive archive)
        {
            var href = node
                .SelectSingleNode(".//draw:image", null)?
                .Attributes["xlink:href"]?.Value;

            if (href == null)
                return "[image]";

            var entry = archive.GetEntry(href.TrimStart('.', '/'));

            if (entry == null)
                return "[missing image]";

            var stream = entry.Open();
            var ms = new MemoryStream();
            stream.CopyTo(ms);

            var base64 = Convert.ToBase64String(ms.ToArray());

            // fallback placeholder (RTF PNG)
            return @"{\pard\plain\fs20 [Image]\par}";
        }
        string ParseSpan(XmlNode node, string styleName)
        {
            var text = EscapeRtf(node.InnerText);

            bool bold = styleName?.Contains("bold") == true;
            bool italic = styleName?.Contains("italic") == true;

            var sb = new StringBuilder();

            if (bold) sb.Append(@"\b ");
            if (italic) sb.Append(@"\i ");

            sb.Append(text);

            if (italic) sb.Append(@"\i0 ");
            if (bold) sb.Append(@"\b0 ");

            return sb.ToString();
        }
        string ParseParagraph(XmlNode p, XmlNamespaceManager ns, ZipArchive archive)
        {
            var sb = new StringBuilder();

            foreach (XmlNode node in p.ChildNodes)
            {
                if (node.Name == "text:span")
                {
                    var style = node.Attributes?["text:style-name"]?.Value;

                    sb.Append(ParseSpan(node, style));
                }
                else if (node.Name == "#text")
                {
                    sb.Append(EscapeRtf(node.InnerText));
                }
                else if (node.Name == "draw:frame")
                {
                    sb.Append(ParseImage(node, archive));
                }
            }

            return sb.ToString();
        }

        async Task LoadOdt(string contentXml, ZipArchive archive, RichEditBox editor)
        {
            var doc = new XmlDocument();
            doc.LoadXml(contentXml);

            var ns = new XmlNamespaceManager(doc.NameTable);
            ns.AddNamespace("text", "urn:oasis:names:tc:opendocument:xmlns:text:1.0");
            ns.AddNamespace("draw", "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0");
            ns.AddNamespace("xlink", "http://www.w3.org/1999/xlink");

            var sb = new StringBuilder();
            sb.Append(@"{\rtf1\ansi");

            foreach (XmlNode p in doc.SelectNodes("//text:p", ns))
            {
                sb.Append(@"\par ");

                sb.Append(ParseParagraph(p, ns, archive));
            }

            sb.Append("}");

            editor.Document.SetText(TextSetOptions.FormatRtf, sb.ToString());
        }

        string GetOdtText(string contentXml)
        {
            var doc = new System.Xml.XmlDocument();
            doc.LoadXml(contentXml);

            var nsManager = new System.Xml.XmlNamespaceManager(doc.NameTable);
            nsManager.AddNamespace("text", "urn:oasis:names:tc:opendocument:xmlns:text:1.0");
            nsManager.AddNamespace("office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0");

            var nodes = doc.SelectNodes("//text:p", nsManager);

            var sb = new StringBuilder();

            foreach (System.Xml.XmlNode node in nodes)
            {
                sb.AppendLine(node.InnerText);
            }

            return sb.ToString();
        }

        protected async override void OnNavigatedTo(NavigationEventArgs e)
        {
            base.OnNavigatedTo(e);

            if (e.Parameter is StorageFile file)
            {
                _currentFile = file;

                string ext = file.FileType.ToLower();

                if (ext == ".rtf")
                {
                    string rtf;

                    using (var stream = await file.OpenStreamForReadAsync())
                    using (var reader = new StreamReader(stream))
                    {
                        rtf = await reader.ReadToEndAsync();
                    }

                    rtf = FixRtf(rtf);

                    Editor.Document.SetText(TextSetOptions.FormatRtf, rtf);
                }

                else if (ext == ".odt")
                {
                    using (var stream = await file.OpenStreamForReadAsync())
                    using (var archive = new ZipArchive(stream, ZipArchiveMode.Read))
                    {
                        var contentEntry = archive.GetEntry("content.xml");
                        if (contentEntry == null) return;

                        string contentXml;

                        using (var s = contentEntry.Open())
                        using (var r = new StreamReader(s))
                            contentXml = await r.ReadToEndAsync();

                        await LoadOdt(contentXml, archive, Editor);
                    }
                }

                else if (ext == ".txt")
                {
                    string text;

                    using (var stream = await file.OpenStreamForReadAsync())
                    using (var reader = new StreamReader(
                        stream,
                        Encoding.UTF8,
                        detectEncodingFromByteOrderMarks: true,
                        bufferSize: 1024))
                    {
                        text = await reader.ReadToEndAsync();
                    }

                    if (text.Contains("�"))
                    {
                        using (var stream = await file.OpenStreamForReadAsync())
                        using (var reader = new StreamReader(stream, Encoding.GetEncoding(1250)))
                        {
                            text = await reader.ReadToEndAsync();
                        }
                    }

                    Editor.Document.SetText(TextSetOptions.None, text);
                }

                fileNameWithPath = file.Path;
                AppTitle.Text = file.Name + " - " + appTitleStr;

                StorageApplicationPermissions.MostRecentlyUsedList.Add(file);
                StorageApplicationPermissions.FutureAccessList.AddOrReplace("CurrentlyOpenFile", file);

                Editor.Document.GetText(TextGetOptions.None, out _lastSavedText);

                ResetDirtyAfterDelay();
            }
        }

        private async void Open_Click(object sender, RoutedEventArgs e)
        {
            FileOpenPicker open = new FileOpenPicker();
            open.SuggestedStartLocation = PickerLocationId.DocumentsLibrary;
            open.FileTypeFilter.Add(".odt");
            open.FileTypeFilter.Add(".rtf");
            open.FileTypeFilter.Add(".txt");

            StorageFile file = await open.PickSingleFileAsync();
            if (file == null) return;

            _currentFile = file;

            string ext = file.FileType.ToLower();

            if (ext == ".rtf")
            {
                string rtf;

                using (var stream = await file.OpenStreamForReadAsync())
                using (var reader = new StreamReader(stream))
                {
                    rtf = await reader.ReadToEndAsync();
                }

                rtf = FixRtf(rtf);

                Editor.Document.SetText(TextSetOptions.FormatRtf, rtf);
            }

            else if (ext == ".odt")
            {
                var dialog = new ContentDialog
                {
                    Title = "Experimental feature",
                    Content = "Support for .odt files is experimental.\n\n" +
                              "Formatting and images may not display correctly.\n" +
                              "The app may become unstable with complex files.\n\n" +
                              "Continue?",
                    PrimaryButtonText = "Continue",
                    CloseButtonText = "Cancel"
                };

                if (await dialog.ShowAsync() != ContentDialogResult.Primary)
                    return;

                using (var stream = await file.OpenStreamForReadAsync())
                using (var archive = new ZipArchive(stream, ZipArchiveMode.Read))
                {
                    var contentEntry = archive.GetEntry("content.xml");
                    if (contentEntry == null) return;

                    string contentXml;

                    using (var s = contentEntry.Open())
                    using (var r = new StreamReader(s))
                        contentXml = await r.ReadToEndAsync();

                    await LoadOdt(contentXml, archive, Editor);
                }
            }

            else if (ext == ".txt")
            {
                string text;

                using (var stream = await file.OpenStreamForReadAsync())
                using (var reader = new StreamReader(
                    stream,
                    Encoding.UTF8,
                    detectEncodingFromByteOrderMarks: true,
                    bufferSize: 1024))
                {
                    text = await reader.ReadToEndAsync();
                }

                if (text.Contains("�"))
                {
                    using (var stream = await file.OpenStreamForReadAsync())
                    using (var reader = new StreamReader(stream, Encoding.GetEncoding(1250)))
                    {
                        text = await reader.ReadToEndAsync();
                    }
                }

                Editor.Document.SetText(TextSetOptions.None, text);
            }

            AppTitle.Text = file.Name + " - " + appTitleStr;
            fileNameWithPath = file.Path;

            StorageApplicationPermissions.MostRecentlyUsedList.Add(file);
            StorageApplicationPermissions.FutureAccessList.AddOrReplace("CurrentlyOpenFile", file);
            Editor.Document.GetText(TextGetOptions.None, out _lastSavedText);
            ResetDirtyAfterDelay();
        }

        private void SubscriptButton_Click(object sender, RoutedEventArgs e)
        {
            Editor.FormatSelected(RichEditHelpers.FormattingMode.Subscript);
        }

        private void SuperScriptButton_Click(object sender, RoutedEventArgs e)
        {
            Editor.FormatSelected(RichEditHelpers.FormattingMode.Superscript);
        }
        private void StrikethroughButton_Click(object sender, RoutedEventArgs e)
        {
            Editor.FormatSelected(RichEditHelpers.FormattingMode.Strikethrough);
        }


        private void NoneNumeral_Click(object sender, RoutedEventArgs e)
        {
            Editor.Document.Selection.ParagraphFormat.ListType = MarkerType.None;
            Editor.Focus(FocusState.Keyboard);
        }

        private void DottedNumeral_Click(object sender, RoutedEventArgs e)
        {
            Editor.Document.Selection.ParagraphFormat.ListType = MarkerType.Bullet;
            Editor.Focus(FocusState.Keyboard);
        }

        private void NumberNumeral_Click(object sender, RoutedEventArgs e)
        {
            Editor.Document.Selection.ParagraphFormat.ListType = MarkerType.Arabic;
            Editor.Focus(FocusState.Keyboard);
        }

        private void LetterSmallNumeral_Click(object sender, RoutedEventArgs e)
        {
            Editor.Document.Selection.ParagraphFormat.ListType = MarkerType.LowercaseEnglishLetter;
            Editor.Focus(FocusState.Keyboard);
        }

        private void LetterBigNumeral_Click(object sender, RoutedEventArgs e)
        {
            Editor.Document.Selection.ParagraphFormat.ListType = MarkerType.UppercaseEnglishLetter;
            Editor.Focus(FocusState.Keyboard);
        }

        private void SmalliNumeral_Click(object sender, RoutedEventArgs e)
        {
            Editor.Document.Selection.ParagraphFormat.ListType = MarkerType.LowercaseRoman;
            Editor.Focus(FocusState.Keyboard);
        }

        private void BigINumeral_Click(object sender, RoutedEventArgs e)
        {
            Editor.Document.Selection.ParagraphFormat.ListType = MarkerType.UppercaseRoman;
            Editor.Focus(FocusState.Keyboard);
        }


        private void AlignRightButton_Click(object sender, RoutedEventArgs e)
        {
            Editor.AlignSelectedTo(RichEditHelpers.AlignMode.Right);
            editor_SelectionChanged(sender, e);
        }

        private void AlignCenterButton_Click(object sender, RoutedEventArgs e)
        {
            Editor.AlignSelectedTo(RichEditHelpers.AlignMode.Center);
            editor_SelectionChanged(sender, e);
        }

        private void AlignLeftButton_Click(object sender, RoutedEventArgs e)
        {
            Editor.AlignSelectedTo(RichEditHelpers.AlignMode.Left);
            editor_SelectionChanged(sender, e);
        }

        private void FindBoxRemoveHighlights()
        {
            ITextRange documentRange = Editor.Document.GetRange(0, TextConstants.MaxUnitCount);
            SolidColorBrush defaultBackground = Editor.Background as SolidColorBrush;
            SolidColorBrush defaultForeground = Editor.Foreground as SolidColorBrush;

            documentRange.CharacterFormat.BackgroundColor = defaultBackground.Color;
            documentRange.CharacterFormat.ForegroundColor = defaultForeground.Color;
        }

        private void RemoveHighlightButton_Click(object sender, RoutedEventArgs e)
        {
            FindBoxRemoveHighlights();
        }

        private void Undo_Click(object sender, RoutedEventArgs e)
        {
            Editor.Document.Undo();
        }

        private void Redo_Click(object sender, RoutedEventArgs e)
        {
            Editor.Document.Redo();
        }
        private void Cut_Click(object sender, RoutedEventArgs e)
        {
            Editor.Document.Selection.Cut();
        }

        private void Copy_Click(object sender, RoutedEventArgs e)
        {
            Editor.Document.Selection.Copy();
        }

        private void Paste_Click(Microsoft.UI.Xaml.Controls.SplitButton sender, Microsoft.UI.Xaml.Controls.SplitButtonClickEventArgs args)
        {
            Editor.Document.Selection.Paste(0);
        }

        private void ZoomSlider_ValueChanged(object sender, Windows.UI.Xaml.Controls.Primitives.RangeBaseValueChangedEventArgs e)
        {
            AnimateZoomSecond(e.OldValue, e.NewValue);
        }


        private void EditorContentHost_SizeChanged(object sender, SizeChangedEventArgs e)
        {

            /*
 
            The status bar is slightly tinted by the mica backdrop

            Clipping the editor is needed, as the editor has a
            shadow. Without the clip, the shadow would be visible
            on the status bar

            */

            RectangleGeometry rectangle = new RectangleGeometry();
            rectangle.Rect = new Rect(0, 0, EditorContentHost.ActualWidth, EditorContentHost.ActualHeight);
            EditorContentHost.Clip = rectangle;
        }

        private void ToggleButton_Checked(object sender, RoutedEventArgs e)
        {

        }
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (Editor != null)
            {
                Editor.Focus(FocusState.Programmatic);

                // Get the position of the last character in the RichEditBox
                int lastPosition = Editor.Document.Selection.EndPosition;

                // Set the selection range to the entire document
                Editor.Document.Selection.SetRange(0, lastPosition);
            }
        }

        private void ToggleButton_Checked_1(object sender, RoutedEventArgs e)
        {
            object value = Editor.Document.Selection.CharacterFormat.Bold = FormatEffect.Toggle;
        }

        private void ToggleButton_Unchecked_1(object sender, RoutedEventArgs e)
        {
            object value = Editor.Document.Selection.CharacterFormat.Bold = FormatEffect.Toggle;
        }

        private async Task<bool?> ShowUnsavedDialog()
        {
            string fileName = AppTitle.Text.Replace(" - " + appTitleStr, "");

            var dialog = new ContentDialog
            {
                Title = "Unsaved changes",
                Content = $"Do you want to save changes to \"{fileName}\"?",
                PrimaryButtonText = "Save",
                SecondaryButtonText = "Don't save",
                CloseButtonText = "Cancel",
                DefaultButton = ContentDialogButton.Primary
            };

            var result = await dialog.ShowAsync();




            if (result == ContentDialogResult.Secondary)
                return true;
            else if(result == ContentDialogResult.Primary)
            {
                await SaveFileAsync(false, "DefaultFull");
                return true;   
            }
            else if (result == ContentDialogResult.Primary && fileName == "Document")
            {
                await SaveFileAsync(true, "DefaultFull");
                return true;
            }
            else if (result == ContentDialogResult.Primary && fileName == "Dokument")
            {
                await SaveFileAsync(true, "DefaultFull");
                return true;
            }

            return false;     
        }

        private void ToggleButton_Checked_2(object sender, RoutedEventArgs e)
        {
            object value = Editor.Document.Selection.CharacterFormat.Italic = FormatEffect.Toggle;
        }

        private void ToggleButton_Checked_3(object sender, RoutedEventArgs e)
        {
            object value = Editor.Document.Selection.CharacterFormat.Strikethrough = FormatEffect.Toggle;
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            Editor.ChangeFontSize((float)2);
        }

        private void SaveAsButton_Click(object sender, RoutedEventArgs e)
        {
            SaveFileAsync(true, "DefaultFull");
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            SaveFileAsync(false, "DefaultFull");
        }


        private async void AddImageButton_Click(Microsoft.UI.Xaml.Controls.SplitButton sender, Microsoft.UI.Xaml.Controls.SplitButtonClickEventArgs args)
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
            }
        }

        private StorageFile _currentFile = null;

        private async Task SaveFileAsync(bool forceSaveAs, string selectedName)
        {
            StorageFile file = _currentFile;

            if (forceSaveAs || file == null)
            {
                FileSavePicker savePicker = new FileSavePicker
                {
                    SuggestedStartLocation = PickerLocationId.DocumentsLibrary,
                    SuggestedFileName = "New Document"
                };

                if (selectedName == "RTF")
                {
                    savePicker.FileTypeChoices.Add("Rich Text Format", new List<string>() { ".rtf" });
                }
                else if (selectedName == "TXT")
                {
                    savePicker.FileTypeChoices.Add("Text Document", new List<string>() { ".txt" });
                }
                else if (selectedName == "DefaultFull")
                {
                    savePicker.FileTypeChoices.Add("Rich Text Format", new List<string>() { ".rtf" });
                    savePicker.FileTypeChoices.Add("Text Document", new List<string>() { ".txt" });
                }

                file = await savePicker.PickSaveFileAsync();
                if (file == null) return;

                _currentFile = file;
            }

            CachedFileManager.DeferUpdates(file);

            using (var stream = await file.OpenAsync(FileAccessMode.ReadWrite))
            {
                stream.Seek(0);

                switch (file.FileType)
                {
                    case ".txt":
                        Editor.Document.GetText(TextGetOptions.None, out string text);

                        using (var writer = new DataWriter(stream))
                        {
                            writer.UnicodeEncoding = UnicodeEncoding.Utf8;
                            writer.WriteString(text);

                            await writer.StoreAsync();
                            await writer.FlushAsync();
                        }
                        break;

                    case ".rtf":
                        Editor.Document.SaveToStream(TextGetOptions.FormatRtf, stream);
                        break;

                    case ".docx":
                        Editor.Document.SaveToStream(TextGetOptions.FormatRtf, stream);
                        break;
                }
            }

            var status = await CachedFileManager.CompleteUpdatesAsync(file);

            if (status != FileUpdateStatus.Complete)
            {
                await new ContentDialog
                {
                    Title = "Save failed",
                    Content = $"File {file.Name} couldn't be saved.",
                    CloseButtonText = "OK"
                }.ShowAsync();
            }
            _currentFile = file;
            _isDirty = false;

            Editor.Document.GetText(TextGetOptions.None, out _lastSavedText);

            AppTitle.Text = file.Name + " - " + appTitleStr;

            StorageApplicationPermissions.MostRecentlyUsedList.Add(file);
        }

        private void CancelColor_Click(object sender, RoutedEventArgs e)
        {
        }

        private void fontbackgroundcolorsplitbutton_Click(Microsoft.UI.Xaml.Controls.SplitButton sender, Microsoft.UI.Xaml.Controls.SplitButtonClickEventArgs args)
        {
            // If you see this, remind me to look into the splitbutton color applying logic
        }

        private void fontcolorsplitbutton_Click(Microsoft.UI.Xaml.Controls.SplitButton sender, Microsoft.UI.Xaml.Controls.SplitButtonClickEventArgs args)
        {
            // If you see this, remind me to look into the splitbutton color applying logic
        }

        private void RadioButton_Click(object sender, RoutedEventArgs e)
        {

        }

        private void ItalicButton_Click(object sender, RoutedEventArgs e)
        {
            Editor.FormatSelected(RichEditHelpers.FormattingMode.Italic);
        }

        private void BoldButton_Click(object sender, RoutedEventArgs e)
        {
            Editor.FormatSelected(RichEditHelpers.FormattingMode.Bold);
        }

        private void UnderlineButton_Click(object sender, RoutedEventArgs e)
        {
            Editor.FormatSelected(RichEditHelpers.FormattingMode.Underline);
        }

        private void AlignAdjustedButton_Click(object sender, RoutedEventArgs e)
        {

        }

        private async void ParagraphButton_Click(object sender, RoutedEventArgs e)
        {
            pararibbontoolbar.ShowParagraphDialog();
        }

        private void editor_SelectionChanged(object sender, RoutedEventArgs e)
        {

            if (Editor.Document.Selection.CharacterFormat.Size > 0)
            {
                //font size is negative when selection contains multiple font sizes
                //FontSizeBox. = Editor.Document.Selection.CharacterFormat.Size;
            }
            //prevent accidental font changes when selection contains multiple styles
            updateFontFormat = false;
            updateFontFormat = true;
            // Get a reference to the RichEditBox control
            RichEditBox richEditBox = Editor;
        }

        private async void Button_Click_3Async(object sender, RoutedEventArgs e)
        {
                
        }

        private void DecreaseFontSize_Click(object sender, RoutedEventArgs e)
        {
            // Get a reference to the RichEditBox control
            RichEditBox richEditBox = Editor;

            // Decrease the font size of the currently selected text by 2 points
            richEditBox.Document.Selection.CharacterFormat.Size -= 2;
        }

        private void IncreaseFontSize_Click(object sender, RoutedEventArgs e)
        {
            // Get a reference to the RichEditBox control
            RichEditBox richEditBox = Editor;

            // Increase the font size of the currently selected text by 2 points
            richEditBox.Document.Selection.CharacterFormat.Size += 2;
        }

        private async void Button_Click_4Async(object sender, RoutedEventArgs e)
        {
            ContentDialog dialog = new ContentDialog();
            dialog.Title = "Insert current date and time";

            // Create a ListView for the user to select the date format
            ListView listView = new ListView();
            listView.SelectionMode = ListViewSelectionMode.Single;

            // Create a list of date formats to display in the ListView
            List<string> dateFormats = new List<string>();
            dateFormats.Add(DateTime.Now.ToString("dd.M.yyyy"));
            dateFormats.Add(DateTime.Now.ToString("M.dd.yyyy"));
            dateFormats.Add(DateTime.Now.ToString("dd MMM yyyy"));
            dateFormats.Add(DateTime.Now.ToString("dddd, dd MMMM yyyy"));
            dateFormats.Add(DateTime.Now.ToString("dd MMMM yyyy"));
            dateFormats.Add(DateTime.Now.ToString("hh:mm:ss tt"));
            dateFormats.Add(DateTime.Now.ToString("HH:mm:ss"));
            dateFormats.Add(DateTime.Now.ToString("dddd, dd MMMM yyyy, HH:mm:ss"));
            dateFormats.Add(DateTime.Now.ToString("dd MMMM yyyy, HH:mm:ss"));
            dateFormats.Add(DateTime.Now.ToString("MMM dd, yyyy"));

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
                string formattedDate = dateFormats[listView.SelectedIndex];
                Editor.Document.Selection.Text = formattedDate;
            };

            // Add a "Cancel" button to the ContentDialog
            dialog.SecondaryButtonText = "Cancel";

            // Show the ContentDialog
            await dialog.ShowAsync();
        }

        private PrintHelper _printHelper;
        private DataTemplate customPrintTemplate;
        private async void MenuFlyoutItem_Click(object sender, RoutedEventArgs e)
        {
            Editor.RequestedTheme = ElementTheme.Light;
            string value = string.Empty;
            _printHelper = new PrintHelper(EditorMandatoryPrintingGrid);
            var printHelperOptions = new PrintHelperOptions(true);
            printHelperOptions.Orientation = PrintOrientation.Default;
            await _printHelper.ShowPrintUIAsync("Print Document", printHelperOptions, true);
            Editor.RequestedTheme = ElementTheme.Default;
        }
        private void pintpreview_Click(object sender, RoutedEventArgs e)
        {
            ribbongrid.Visibility = Visibility.Collapsed;
            RulerBorder.Visibility = Visibility.Collapsed;
            ZoomStack.Visibility = Visibility.Collapsed;
            Editor.IsEnabled = false;
            PrintPreviewRibbon.Visibility = Visibility.Visible;
        }
        private void closeprintpreviewclick(object sender, RoutedEventArgs e)
        {
            ribbongrid.Visibility = Visibility.Visible;
            RulerBorder.Visibility = Visibility.Visible;
            ZoomStack.Visibility = Visibility.Visible;
            Editor.IsEnabled = true;
            PrintPreviewRibbon.Visibility = Visibility.Collapsed;
        }

        bool isTextChanged = false;
        private readonly bool isCopy;

        private async void OnCloseRequest(object sender, SystemNavigationCloseRequestedPreviewEventArgs e)
        {
            if (_isDirty == true)
            {
                e.Handled = true;

                var result = await ShowUnsavedDialog();
                if (result == true)
                {
                    ApplicationView.GetForCurrentView().TryConsolidateAsync();
                }
            }
        }

        private async void SaveAsRTF_Click(object sender, RoutedEventArgs e)
        {
            SaveFileAsync(true, "RTF");
        }

        private async void SaveAsDOCX_Click(object sender, RoutedEventArgs e)
        {
            SaveFileAsync(true, "DefaultFull");
            // Should be changed from DefaultFull to DOCX once .docx save support is added
        }

        private async void SaveAsODT_Click(object sender, RoutedEventArgs e)
        {
            SaveFileAsync(true, "DefaultFull");
            // Should be changed from DefaultFull to ODT once .odt save support is added
        }

        private async void SaveAsTXT_Click(object sender, RoutedEventArgs e)
        {
            SaveFileAsync(true, "TXT");
        }

        private void DisableDocTree()
        {
            DocTreeBorder.Visibility = Visibility.Collapsed;
            Grid.SetColumn(EditorContentHost,0);
            Grid.SetColumn(RulerBorder, 0);
        }

        private void EnableDocTree()
        {
            DocTreeBorder.Visibility = Visibility.Visible;
            Grid.SetColumn(EditorContentHost, 1);
            Grid.SetColumn(RulerBorder, 1);
        }

        private async void SaveAsOther_Click(object sender, RoutedEventArgs e)
        {
            SaveFileAsync(true, "DefaultFull");
        }

        private async void NewDoc_Click(object sender, RoutedEventArgs e)
        {
            if (_isDirty)
            {
                var proceed = await ShowUnsavedDialog();
                if (proceed != true)
                    return;
            }

            Editor.Document.SetText(TextSetOptions.None, string.Empty);
            _currentFilePath = null;
            MarkClean();
        }

        private async void SendButton_Click(object sender, RoutedEventArgs e)
        {
            DataTransferManager.ShowShareUI();
        }

        private void AlignJustifyButton_Click(object sender, RoutedEventArgs e)
        {
            Editor.AlignSelectedTo(RichEditHelpers.AlignMode.Justify);
            editor_SelectionChanged(sender, e);
        }

        private void DecreaseZoomButton_Click(object sender, RoutedEventArgs e)
        {
            ZoomSlider.Value = Math.Max(ZoomSlider.Value - 0.1, ZoomSlider.Minimum);
        }

        private void IncreaseZoomButton_Click(object sender, RoutedEventArgs e)
        {
            ZoomSlider.Value = Math.Min(ZoomSlider.Value + 0.1, ZoomSlider.Maximum);
        }

        private void AnimateZoomSecond(double fromValue, double toValue)
        {
            RichTextScrollView.ChangeView(0, 0, (float)ZoomSlider.Value);
            float zoom = (float)ZoomSlider.Value;
            TextRuler.Width = 814 * zoom;
        }

        private void MenuCut_Click(object sender, RoutedEventArgs e)
        {
            Editor.Document.Selection.Cut();
        }

        private void MenuCopy_Click(object sender, RoutedEventArgs e)
        {
            Editor.Document.Selection.Copy();
        }

        private void MenuPaste_Click(object sender, RoutedEventArgs e)
        {
            Editor.Document.Selection.Paste(0);
        }

        private async void MenuParagraph_Click(object sender, RoutedEventArgs e)
        {
            pararibbontoolbar.ShowParagraphDialog();
        }

        public string GetText(RichEditBox RichEditor)
        {
            RichEditor.Document.GetText(TextGetOptions.FormatRtf, out string Text);
            ITextRange Range = RichEditor.Document.GetRange(0, Text.Length);
            Range.GetText(TextGetOptions.FormatRtf, out string Value);
            return Value;
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            baseRulerWidth = TextRuler.ActualWidth;
            _isDirty = false;
        }

        private void SetParagraphIndents(float leftIndent, float rightIndent, float firstLineIndent, bool applyToSelectionOnly = true)
        {
            // Get the ITextDocument interface for the RichEditBox's document
            ITextDocument document = Editor.Document;

            // Get the current selection's start and end positions
            int start = document.Selection.StartPosition;
            int end = document.Selection.EndPosition;

            // If applyToSelectionOnly is true, check if there's any selected text in the RichEditBox
            if (applyToSelectionOnly && start == end)
            {
                //return;
            }

            // Get the ITextRange interface for the selection or the entire document
            ITextRange textRange;
            if (applyToSelectionOnly)
            {
                textRange = document.Selection;
            }
            else
            {
                textRange = document.GetRange(0, GetText(Editor).Length);
            }

            // Get the ITextParagraphFormat interface for the text range
            ITextParagraphFormat paragraphFormat = textRange.ParagraphFormat;

            // Set the left and right indents for the current selection's paragraph(s)
            try
            {
                if (document.Selection.Length != 0)
                {
                    paragraphFormat.SetIndents(firstLineIndent, leftIndent, rightIndent);
                }
                else
                {
                    document.GetRange(document.Selection.StartPosition, document.Selection.EndPosition + 1);
                    paragraphFormat.SetIndents(firstLineIndent, leftIndent, rightIndent);
                }
            }
            catch
            {

            }

            // Apply the new paragraph format to the current selection or the entire document
            textRange.ParagraphFormat = paragraphFormat;

            // LeftIndent.Text = leftIndent.ToString();

            // RightIndent.Text = rightIndent.ToString();
        }

        

        private void MenuFlyoutItem_Click_1(object sender, RoutedEventArgs e)
        {

        }

        private void PageSetup_Click(object sender, RoutedEventArgs e)
        {
            openpageprop();
        }

        private async void opennotimplement() 
        {
            // Create an instance of the ParagraphDialog
            NoImplement noimple = new NoImplement();

            // Show the dialog and wait for the user's input
            ContentDialogResult result = await noimple.ShowAsync();
        }

        private async void openpageprop()
        {
            // Create an instance of the ParagraphDialog
            Pageprop pageprop = new Pageprop();

            // Show the dialog and wait for the user's input
            ContentDialogResult result = await pageprop.ShowAsync();

            // If the user clicked the OK button, adjust the properties of the RichEditBox
            if (result == ContentDialogResult.Primary)
            {

                // Get the values from the dialog's TextBoxes and ComboBoxes
                TextBox LeftMarginTextBox = (TextBox)pageprop.FindName("LeftMarginTextBox");
                TextBox RightMarginTextBox = (TextBox)pageprop.FindName("RightMarginTextBox");
                TextBox TopMarginTextBox = (TextBox)pageprop.FindName("TopMarginTextBox");
                TextBox BottomMarginTextBox = (TextBox)pageprop.FindName("BottomMarginTextBox");

                TextBlock marginsname = (TextBlock)pageprop.FindName("marginsname");

                ComboBox PaperTypeCombo = (ComboBox)pageprop.FindName("PaperTypeCombo");
                RadioButton orientationportait = (RadioButton)pageprop.FindName("orientationportait");
                CheckBox printpagenumbers = (CheckBox)pageprop.FindName("printpagenumbers");

                // Save the selected paper size and orientation
                var settings = ApplicationData.Current.LocalSettings;
                if (PaperTypeCombo.SelectedItem != null)
                {
                    string selectedPaperSize = (PaperTypeCombo.SelectedItem as ComboBoxItem).Content.ToString();
                    settings.Values["papersize"] = selectedPaperSize;
                }

                settings.Values["orientation"] = orientationportait.IsChecked == true ? "Portrait" : "Landscape";

                // Save margin values
                // settings.Values["pagesetupLmargin"] = unitConverter.ConvertToUnit(double.Parse(LeftMarginTextBox.Text), marginsname.Text);
                // settings.Values["pagesetupRmargin"] = unitConverter.ConvertToUnit(double.Parse(RightMarginTextBox.Text), marginsname.Text);
                // settings.Values["pagesetupTmargin"] = unitConverter.ConvertToUnit(double.Parse(TopMarginTextBox.Text), marginsname.Text);
                // settings.Values["pagesetupBmargin"] = unitConverter.ConvertToUnit(double.Parse(BottomMarginTextBox.Text), marginsname.Text);

                // Save Print Page Numbers setting
                settings.Values["isprintpagenumbers"] = printpagenumbers.IsChecked == true ? "yes" : "no";

                Dictionary<string, (double Width, double Height)> paperSizes = pageprop.paperSizes;

                string selectedPaperSizea = (PaperTypeCombo.SelectedItem as ComboBoxItem)?.Content?.ToString();
                if (!string.IsNullOrEmpty(selectedPaperSizea) && paperSizes.TryGetValue(selectedPaperSizea, out var dimensions))
                {
                    double originalWidth = 812; // Original RichEditBox width
                    double originalHeight = 1116; // Original RichEditBox height

                    double width = dimensions.Width;
                    double height = dimensions.Height;

                    // Calculate the scaling factors for width and height to maintain the aspect ratio
                    double widthScaleFactor = width / originalWidth;
                    double heightScaleFactor = height / originalHeight;

                    // Determine the scaling factor that fits the width within the original width
                    double widthFitScaleFactor = originalWidth / width;

                    // Determine the scaling factor that fits the height within the original height
                    double heightFitScaleFactor = originalHeight / height;

                    // Choose the minimum scaling factor to ensure the content fits entirely within the original dimensions
                    double minScaleFactor = Math.Min(widthFitScaleFactor, heightFitScaleFactor);

                    // Apply the minimum scaling factor to both width and height to maintain the aspect ratio
                    width *= minScaleFactor;
                    height *= minScaleFactor;

                    // Set the UWP's RichEditBox width and height
                    // EditorGrid.Width = width;
                    // EditorGrid.Height = height;
                }
            }
        }

        private void PageSetupMenuItem_Click(object sender, RoutedEventArgs e)
        {
            openpageprop();
        }

        private void Editor_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            // RulerBorder.Width = Editor.Width;
        }

        private void SelectWord(RichEditBox box, bool forward)
        {
            var selection = box.Document.Selection;

            if (forward)
            {
                selection.MoveEnd(TextRangeUnit.Word, 1);
            }
            else
            {
                selection.MoveStart(TextRangeUnit.Word, -1);
            }
        }

        private async void PrintPreviewPrintButton_Click(object sender, RoutedEventArgs e)
        {
            Editor.RequestedTheme = ElementTheme.Light;
            string value = string.Empty;
            _printHelper = new PrintHelper(EditorMandatoryPrintingGrid);
            var printHelperOptions = new PrintHelperOptions(true);
            printHelperOptions.Orientation = PrintOrientation.Default;
            await _printHelper.ShowPrintUIAsync("Print Document", printHelperOptions, true);
            Editor.RequestedTheme = ElementTheme.Default;
        }

        private async void Button_Click_2(object sender, RoutedEventArgs e)
        {
            await Windows.System.Launcher.LaunchUriAsync(new Uri("https://github.com/Lixkote/RectifyPad"));
        }

        private void ShareButton_Click(object sender, RoutedEventArgs e)
        {
            DataTransferManager.ShowShareUI();
        }
        async void DataTransferManager_DataRequested(DataTransferManager sender, DataRequestedEventArgs args)
        {
            DataRequest request = args.Request;
            request.Data.Properties.Title = "My Custom Subject";

            // Retrieve the RTF content from the RichEditBox.
            string rtfContent;
            Editor.Document.GetText(Windows.UI.Text.TextGetOptions.FormatRtf, out rtfContent);

            // Access the temporary folder.
            var storageFolder = Windows.Storage.ApplicationData.Current.TemporaryFolder;
            var fileName = "Document.rtf";

            // Create a new file.
            var rtfFile = await storageFolder.CreateFileAsync(fileName, Windows.Storage.CreationCollisionOption.ReplaceExisting);

            // Write the RTF content to the new file.
            await Windows.Storage.FileIO.WriteTextAsync(rtfFile, rtfContent);

            // Attach the file to the DataRequest.
            request.Data.SetStorageItems(new List<Windows.Storage.IStorageItem> { rtfFile });
        }


        private void SettingsButton_Click(object sender, RoutedEventArgs e)
        {
            this.Frame.Navigate(typeof(SettingsPage));
        }

        private void QuickPrint_Click(object sender, RoutedEventArgs e)
        {

        }

        private void CloseMenuFlyoutItem_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Exit();
        }

        private void Editor_Loaded(object sender, RoutedEventArgs e)
        {
            var format = Editor.Document.GetDefaultParagraphFormat();
            format.ListStart = 1;
            Editor.Document.SetDefaultParagraphFormat(format);
        }

        private int FindWordStart(RichEditBox richEditBox, int position)
        {
            string text;
            richEditBox.Document.GetText(Windows.UI.Text.TextGetOptions.None, out text);
            int start = position;
            while (start > 0 && !char.IsWhiteSpace(text[start - 1]) && !IsPunctuation(text[start - 1]))
            {
                start--;
            }
            return start;
        }

        private int FindWordEnd(RichEditBox richEditBox, int position)
        {
            string text;
            richEditBox.Document.GetText(Windows.UI.Text.TextGetOptions.None, out text);
            int end = position;
            while (end < text.Length && !char.IsWhiteSpace(text[end]) && !IsPunctuation(text[end]))
            {
                end++;
            }
            return end;
        }

        private void ShadowRect_Loaded(object sender, RoutedEventArgs e)
        {
            shadow.Receivers.Add(DocTree);
        }

        private void ShadowRectR_Loaded(object sender, RoutedEventArgs e)
        {
            shadow.Receivers.Add(TextRuler);
        }


        private bool IsPunctuation(char c)
        {
            return char.IsPunctuation(c);
        }

        private void Editor_KeyDown_1(object sender, KeyRoutedEventArgs e)
        {
            RichEditBox richEditBox = sender as RichEditBox;
            bool ctrl = Window.Current.CoreWindow.GetKeyState(VirtualKey.Control).HasFlag(CoreVirtualKeyStates.Down);
            bool shift = Window.Current.CoreWindow.GetKeyState(VirtualKey.Shift).HasFlag(CoreVirtualKeyStates.Down);
            
            if (e.Key == VirtualKey.Tab)
            {
                if (richEditBox != null)
                {
                    richEditBox.Document.Selection.TypeText("\t");
                    e.Handled = true;
                }
            }
            else if (e.Key == VirtualKey.A && (Window.Current.CoreWindow.GetKeyState(VirtualKey.Control) & CoreVirtualKeyStates.Down) == CoreVirtualKeyStates.Down)
            {
                if (richEditBox != null)
                {
                    if (Editor != null)
                    {
                        string text;
                        Editor.Document.GetText(TextGetOptions.None, out text);
                        Editor.Focus(FocusState.Programmatic);
                        int lastPosition = Editor.Document.Selection.EndPosition;

                        Editor.Document.Selection.SetRange(0, text.Length);
                    }
                    e.Handled = true;
                }
            }
            else if (e.Key == VirtualKey.Back && (Window.Current.CoreWindow.GetKeyState(VirtualKey.Control) & CoreVirtualKeyStates.Down) == CoreVirtualKeyStates.Down)
            {
                var selection = Editor.Document.Selection;
                var startPos = selection.StartPosition;
                int wordStart = FindWordStart(Editor, startPos);
                selection.SetRange(wordStart, startPos);
                selection.Delete(Windows.UI.Text.TextRangeUnit.Character, 1);
                e.Handled = true;
            }
            else if (e.Key == VirtualKey.Delete && (Window.Current.CoreWindow.GetKeyState(VirtualKey.Control) & CoreVirtualKeyStates.Down) == CoreVirtualKeyStates.Down)
            {
                var selection = Editor.Document.Selection;
                var endPos = selection.EndPosition;
                int wordEnd = FindWordEnd(Editor, endPos);
                selection.SetRange(endPos, wordEnd);
                selection.Delete(Windows.UI.Text.TextRangeUnit.Character, 1);

                e.Handled = true;
            }

            if (ctrl && shift)
            {
                if (e.Key == VirtualKey.Left)
                {
                    SelectWord(richEditBox, false);
                    e.Handled = true;
                }
                else if (e.Key == VirtualKey.Right)
                {
                    SelectWord(richEditBox, true);
                    e.Handled = true;
                }
            }
        }

        private void DocsTreeToggle_Click(object sender, RoutedEventArgs e)
        {
            if (DocsTreeToggle.IsChecked)
            {
                EnableDocTree();
            }
            else
            {
                DisableDocTree();
            }
        }
    }
}
