using DocumentFormat.OpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.System;
using Windows.UI.Text;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Navigation;

// Szablon elementu Kontrolka użytkownika jest udokumentowany na stronie https://go.microsoft.com/fwlink/?LinkId=234236

namespace WordPad.WordPadUI
{
    public sealed partial class TextRuler : UserControl
    {
        public RichEditBox editor;
        public ScrollViewer editorScroll;
        bool isTextSelectionChanging;

        public Slider ZoomSlider { get; private set; }

        public TextRuler()
        {
            this.InitializeComponent();
        }
        public string GetText(RichEditBox RichEditor)
        {
            RichEditor.Document.GetText(TextGetOptions.FormatRtf, out string Text);
            ITextRange Range = RichEditor.Document.GetRange(0, Text.Length);
            Range.GetText(TextGetOptions.FormatRtf, out string Value);
            return Value;
        }

        private void SCR3_ViewChanged(object sender, ScrollViewerViewChangedEventArgs e)
        {
            ZoomSlider.Value = editorScroll.ZoomFactor;
        }

        private void TabIndent_ValueChanged(object sender, Windows.UI.Xaml.Controls.Primitives.RangeBaseValueChangedEventArgs e)
        {
            if (isTextSelectionChanging == false) SetParagraphIndents((float)LeftInd.Value, (float)RightInd.Value, (float)TabIndent.Value, true);
        }

        private void SetParagraphIndents(float leftIndent, float rightIndent, float firstLineIndent, bool applyToSelectionOnly = true)
        {
            // Get the ITextDocument interface for the RichEditBox's document
            ITextDocument document = editor.Document;

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
                textRange = document.GetRange(0, GetText(editor).Length);
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
        private void LeftInd_ValueChanged(object Sender, Windows.UI.Xaml.Controls.Primitives.RangeBaseValueChangedEventArgs EvArgs)
        {
            if (isTextSelectionChanging == false) SetParagraphIndents((float)LeftInd.Value, (float)RightInd.Value, (float)TabIndent.Value, true);
        }

        private void RightInd_ValueChanged(object Sender, Windows.UI.Xaml.Controls.Primitives.RangeBaseValueChangedEventArgs EvArgs)
        {
            if (isTextSelectionChanging == false) SetParagraphIndents((float)LeftInd.Value, (float)RightInd.Value, (float)TabIndent.Value, true);
        }

        private void Border_Loaded(object sender, RoutedEventArgs e)
        {
            // shadow.Receivers.Add(test);
        }
    }
}
