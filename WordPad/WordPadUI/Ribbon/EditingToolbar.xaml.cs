using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.UI;
using Windows.UI.Text;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Navigation;
using WordPad.Helpers;

// Szablon elementu Kontrolka użytkownika jest udokumentowany na stronie https://go.microsoft.com/fwlink/?LinkId=234236

namespace WordPad.WordPadUI.Ribbon
{
    public sealed partial class EditingToolbar : UserControl
    {
        public RichEditBox Editor { get; set; }
        public EditingToolbar()
        {
            this.InitializeComponent();
        }

        private void FindButton_Click(object sender, RoutedEventArgs e)
        {
            FindBoxHighlightMatches();
        }

        private void FindBoxHighlightMatches()
        {

            FindBoxRemoveHighlights();

            Color highlightBackgroundColor = (Color)Application.Current.Resources["SystemColorHighlightColor"];
            Color highlightForegroundColor = (Color)Application.Current.Resources["SystemColorHighlightTextColor"];

            string textToFind = findBox.Text;
            if (textToFind != null)
            {
                ITextRange searchRange = Editor.Document.GetRange(0, 0);
                while (searchRange.FindText(textToFind, TextConstants.MaxUnitCount, FindOptions.None) > 0)
                {
                    searchRange.CharacterFormat.BackgroundColor = highlightBackgroundColor;
                    searchRange.CharacterFormat.ForegroundColor = highlightForegroundColor;
                }
            }
        }

        private void FindBoxRemoveHighlights()
        {
            ITextRange documentRange = Editor.Document.GetRange(0, TextConstants.MaxUnitCount);
            SolidColorBrush defaultBackground = Editor.Background as SolidColorBrush;
            SolidColorBrush defaultForeground = Editor.Foreground as SolidColorBrush;

            documentRange.CharacterFormat.BackgroundColor = defaultBackground.Color;
            documentRange.CharacterFormat.ForegroundColor = defaultForeground.Color;
        }

        private void ReplaceButton_Click(object sender, RoutedEventArgs e)
        {

        }

        private void SelectAllButton_Click(object sender, RoutedEventArgs e)
        {
            if (Editor != null)
            {
                string text;
                Editor.Document.GetText(TextGetOptions.None, out text);
                Editor.Focus(FocusState.Programmatic);
                int lastPosition = Editor.Document.Selection.EndPosition;

                Editor.Document.Selection.SetRange(0, text.Length);
            }
        }


        /// The button click events defined below are click events for the flyouts, not the main ribbon toolbar's buttons.


        private void ReplaceSelected_Click(object sender, RoutedEventArgs e)
        {
            Editor.Replace(false, replaceBox.Text);
        }

        private void ReplaceAll_Click(object sender, RoutedEventArgs e)
        {
            Editor.Replace(true, find: findBox.Text, replace: replaceBox.Text);
        }

        private void RemoveHighlightButton_Click(object sender, RoutedEventArgs e)
        {
            FindBoxRemoveHighlights();
        }
    }
}
