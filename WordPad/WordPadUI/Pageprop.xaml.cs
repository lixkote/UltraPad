using RectifyPad;
using RectifyPad;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Windows.Storage;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Navigation;

namespace WordPad.WordPadUI
{
    public sealed partial class Pageprop : ContentDialog
    {
        public Dictionary<string, (double Width, double Height)> paperSizes;
        private Thickness currentMargins;
        private bool isOrientationChanged = false;
        public Pageprop()
        {
            InitializeComponent();
            LoadUI();
        }

        private void LoadUI()
        {
            var settings = ApplicationData.Current.LocalSettings;

            // Define the dimensions for each paper size based on A4 as a reference
            paperSizes = new Dictionary<string, (double Width, double Height)>
            {
                { "A3", (165 * 297 / 210, 165 * 210 / 297) },
                { "A4", (165, 120) },
                { "B5_1", (120 * 176 / 250, 165 * 250 / 176) },
                { "B5_2", (120 * 176 / 250, 165 * 250 / 176) },
                { "Executive", (120 * 184 / 215, 165 * 215 / 184) },
                { "Legal", (120 * 216 / 356, 165 * 356 / 216) },
                { "Letter", (120 * 216 / 279, 165 * 279 / 216) },
                { "Tabloid", (120 * 279 / 432, 165 * 432 / 279) }
            };

            if (settings.Values.TryGetValue("papersize", out object value))
            {
                string orientation = (string)settings.Values["orientation"];
                string selectedPaperSize = value.ToString();

                if (paperSizes.TryGetValue(selectedPaperSize, out var dimensions))
                {
                    SetPaperSizeAndOrientation(dimensions.Width, dimensions.Height, orientation);
                    PaperTypeCombo.SelectedIndex = paperSizes.Keys.ToList().IndexOf(selectedPaperSize);
                }
            }

            // Load margin values and set limits
            LoadMarginValues();

            // Load the Print Page Numbers setting
            LoadPrintPageNumbersSetting();

            // Load the saved unit value
            string unit = (string)Windows.Storage.ApplicationData.Current.LocalSettings.Values["unit"];
            if (unit == "inches")
            {
                marginsnameinch.Visibility = Visibility.Visible;
            }
            if (unit == "cm")
            {
                marginsnamecm.Visibility = Visibility.Visible;
            }
            if (unit == "points")
            {
                marginsnamept.Visibility = Visibility.Visible;
            }
            if (unit == "picas")
            {
                marginsnamepicas.Visibility = Visibility.Visible;
            }
        }

        private void SetPaperSizeAndOrientation(double width, double height, string orientation)
        {
            // Swap width and height if orientation is Landscape
            if (orientation == "Portrait")
            {
                double temp = width;
                width = height;
                height = temp;
                isOrientationChanged = true;
            }

            Paper.Width = width;
            Paper.Height = height;

            // Update the exampletextgrid margin
            //UpdateMarginPreview();

            // Update the orientation radio buttons
            orientationportait.IsChecked = orientation == "Landscape";
            orientationlandscape.IsChecked = orientation == "Portrait";
        }

        private void LoadMarginValues()
        {
            // Load margin values and set limits
            string unit = (string)Windows.Storage.ApplicationData.Current.LocalSettings.Values["unitSetting"];
            double maxMargin = 100.0; // Set a maximum margin value

            LoadAndLimitMarginValue(LeftMarginTextBox, "pagesetupLmargin", unit, maxMargin);
            LoadAndLimitMarginValue(RightMarginTextBox, "pagesetupRmargin", unit, maxMargin);
            LoadAndLimitMarginValue(TopMarginTextBox, "pagesetupTmargin", unit, maxMargin);
            LoadAndLimitMarginValue(BottomMarginTextBox, "pagesetupBmargin", unit, maxMargin);

            // Update the exampletextgrid margin based on the loaded values
            //UpdateMarginPreview();
        }

        private void LoadAndLimitMarginValue(TextBox textBox, string settingKey, string unit, double maxMargin)
        {
            var settings = ApplicationData.Current.LocalSettings;

            if (settings.Values.TryGetValue(settingKey, out object value))
            {
                // double margin = ConvertFromUnit(Convert.ToDouble(value), "Inches");
                Debug.WriteLine("unitsetting!" + value);

                // Limit the margin value
                // margin = Math.Min(maxMargin, margin);

                //string formattedMargin = ConvertToUnit(margin, unit).ToString("0.##"); // Display with up to 2 decimal places
                //textBox.Text = formattedMargin.TrimEnd('0').TrimEnd('.'); // Remove trailing "00" or "." if present
            }
        }

        private void UpdateMarginPreview()
        {
            try
            {
                double left = double.Parse(LeftMarginTextBox.Text);
                double right = double.Parse(RightMarginTextBox.Text);
                double top = double.Parse(TopMarginTextBox.Text);
                double bottom = double.Parse(BottomMarginTextBox.Text);

                string unit = (string)Windows.Storage.ApplicationData.Current.LocalSettings.Values["unitSetting"];

                double maxWidth = exampletextgrid.Width;
                double maxHeight = exampletextgrid.Height;

                double maxLeftMargin = maxWidth - ConvertToUnit(right, unit);
                double maxRightMargin = maxWidth - ConvertToUnit(left, unit);
                double maxTopMargin = maxHeight - ConvertToUnit(bottom, unit);
                double maxBottomMargin = maxHeight - ConvertToUnit(top, unit);

                double finalLeftMargin = Math.Max(0, Math.Min(ConvertToUnit(left, unit), maxLeftMargin));
                double finalRightMargin = Math.Max(0, Math.Min(ConvertToUnit(right, unit), maxRightMargin));
                double finalTopMargin = Math.Max(0, Math.Min(ConvertToUnit(top, unit), maxTopMargin));
                double finalBottomMargin = Math.Max(0, Math.Min(ConvertToUnit(bottom, unit), maxBottomMargin));

                currentMargins = new Thickness(
                    finalLeftMargin,
                    finalTopMargin,
                    finalRightMargin,
                    finalBottomMargin);

                exampletextgrid.Margin = currentMargins;
                Debug.WriteLine(currentMargins.ToString());
                Debug.WriteLine(maxLeftMargin.ToString());
            }
            catch (Exception ex)
            {
                // Handle any exceptions that might occur during parsing or margin calculation
                Debug.WriteLine("An error occurred: " + ex.Message);
            }
        }

        private void LoadPrintPageNumbersSetting()
        {
            string no = "no";
            var settings = ApplicationData.Current.LocalSettings;
            if (settings.Values.TryGetValue("isprintpagenumbers", out object value))
            {
                string yesorno = value.ToString();
                if (yesorno != null)
                {
                    if (yesorno != no)
                    {
                        printpagenumbers.IsChecked = true;
                    }
                }
            }
        }

        private void orientationportait_Checked(object sender, RoutedEventArgs e)
        {
            if (isOrientationChanged)
            {
                // Swap width and height back to Portrait orientation
                double temp = Paper.Width;
                Paper.Width = Paper.Height;
                Paper.Height = temp;
                isOrientationChanged = false;
            }

            //UpdateMarginPreview();
        }

        private void orientationlandscape_Checked(object sender, RoutedEventArgs e)
        {
            if (!isOrientationChanged)
            {
                // Swap width and height for Landscape orientation
                double temp = Paper.Width;
                Paper.Width = Paper.Height;
                Paper.Height = temp;
                isOrientationChanged = true;
            }

            //UpdateMarginPreview();
        }

        private double ConvertToUnit(double value, string unit)
        {
            switch (unit)
            {
                case "Inches":
                    return value;
                case "Centimeters":
                    return value * 2.54; // Convert inches to centimeters
                case "Points":
                    return value * 72; // Convert inches to points
                case "Picas":
                    return value * 6; // Convert inches to picas
                default:
                    return value; // Default to inches
            }
        }

        private double ConvertFromUnit(double value, string unit)
        {
            switch (unit)
            {
                case "Inches":
                    return value;
                case "Centimeters":
                    return value / 2.54; // Convert centimeters to inches
                case "Points":
                    return value / 72; // Convert points to inches
                case "Picas":
                    return value / 6; // Convert picas to inches
                default:
                    return value; // Default to inches
            }
        }

        private void PaperTypeCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (PaperTypeCombo.SelectedItem != null)
            {
                string selectedPaperSize = (PaperTypeCombo.SelectedItem as ComboBoxItem).Content.ToString();
                if (paperSizes.TryGetValue(selectedPaperSize, out var dimensions))
                {
                    double width = dimensions.Width;
                    double height = dimensions.Height;

                    // Calculate the new width and height based on proportions
                    double originalWidth = 812; // Original RichEditBox width
                    double originalHeight = 1151; // Original RichEditBox height

                    if (width > height)
                    {
                        // Landscape paper size, adjust height while maintaining proportions
                        double scaleFactor = height / originalHeight;
                        height = dimensions.Height;
                        width = originalWidth * scaleFactor;
                    }
                    else
                    {
                        // Portrait paper size, adjust width while maintaining proportions
                        double scaleFactor = width / originalWidth;
                        width = dimensions.Width;
                        height = originalHeight * scaleFactor;
                    }

                    // Set the size of the Paper grid
                    Paper.Width = width;
                    Paper.Height = height;
                }
            }
        }
    }
}
