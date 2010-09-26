using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using Microsoft.Office.Interop.PowerPoint;
using System.Windows.Forms;

namespace PowerPointLaTeX.InlineFormulas
{
    static class Alignment
    {
        public static void PrepareTextRange(Shape picture, TextRange codeRange)
        {
            AdaptFontSize(picture, codeRange);

            // disable word wrap
            codeRange.ParagraphFormat.WordWrap = Microsoft.Office.Core.MsoTriState.msoFalse;

            // fill the text up with none breaking space to make it "wrap around" the formula
            FillTextRange(codeRange, picture.Width);
        }

        private static void AdaptFontSize(Shape picture, TextRange codeRange)
        {
            float baselineOffset = picture.LaTeXTags().BaseLineOffset;
            float heightInPts = DPIHelper.PixelsPerEmHeightToFontSize(picture.Height);

            FontFamily fontFamily = GetFontFamily(codeRange);

            // from top to baseline
            float ascentHeight = (float)(codeRange.Font.Size * ((float)fontFamily.GetCellAscent(FontStyle.Regular) / fontFamily.GetEmHeight(FontStyle.Regular)));
            float descentRatio = (float)fontFamily.GetCellDescent(FontStyle.Regular) / fontFamily.GetEmHeight(FontStyle.Regular);
            float descentHeight = (float)(codeRange.Font.Size * descentRatio);

            float ascentSize = Math.Max(0, 1 - baselineOffset) * heightInPts;
            float descentSize = Math.Max(0, baselineOffset) * heightInPts;

            float factor = 1.0f;
            if (ascentSize > 0)
            {
                factor = Math.Max(factor, ascentSize / ascentHeight);
            }
            if (descentSize > 0)
            {
                factor = Math.Max(factor, descentSize / descentHeight);
            }

            if (factor <= 1.5f)
            {
                codeRange.Font.Size *= factor;
            }
            else
            {
                // don't let it get too big (starts to look ridiculous otherwise)

                // keep linespacing intact (assuming that the line spacing scales with the font height)
                float lineSpacing = (float)(codeRange.Font.Size * ((float)fontFamily.GetLineSpacing(FontStyle.Regular) / fontFamily.GetEmHeight(FontStyle.Regular)));
                // additional line spacing
                // TODO: figure out what the 5.0f is used for and extract it into a constant!.. >_< [9/26/2010 Andreas]
                codeRange.Font.Size *= (lineSpacing + 5.0f - codeRange.Font.Size + heightInPts) / lineSpacing;
                // just ignore the baseline offset
                picture.LaTeXTags().BaseLineOffset.value = descentRatio;
            }
        }

        private static void FillTextRange(TextRange range, float minWidth)
        {
            const char ThinSpace = (char)8201;
            // space that doesn't allow a line break in the middle
            const char NoneBreakingSpace = (char)160;
            // new line that doesn't begin a new paragraph
            const char LineSeparator = (char)8232;

            string fillUnit = ThinSpace.ToString() + NoneBreakingSpace.ToString();
            
            range.Text = fillUnit;

            // line-breaks are futile, so stop filling if one happens 
            float oldHeight = range.BoundHeight;
            while (range.BoundWidth < minWidth && oldHeight == range.BoundHeight)
            {
                range.Text += fillUnit;
            }
            if (oldHeight != range.BoundHeight)
            {
                range.Text = range.Text.Remove(range.Text.Length - 2, 2);
                range.Text += LineSeparator.ToString(); 
            }
        }

        public static void Align(TextRange codeRange, Shape picture)
        {
            // interesting fact: text filled with (at most one line of none-breaking) spaces -> BoundHeight == EmSize
            float fontHeight = codeRange.BoundHeight;
            FontFamily fontFamily = GetFontFamily(codeRange);
            // from top to baseline
            float baselineHeight = (float)(fontHeight * ((float)fontFamily.GetCellAscent(FontStyle.Regular) / fontFamily.GetLineSpacing(FontStyle.Regular)));

            picture.Left = codeRange.BoundLeft;
            picture.Top = codeRange.BoundTop + baselineHeight - (1.0f - picture.LaTeXTags().BaseLineOffset) * picture.Height;

        }

        private static FontFamily GetFontFamily(TextRange textRange)
        {
            FontFamily fontFamily;
            try
            {
                fontFamily = new FontFamily(textRange.Font.Name);
            }
            catch (Exception exception)
            {
                // TODO: add message box and inform the user about it [9/20/2010 Andreas]
                MessageBox.Show("Failed to load font information (using Times New Roman as substitute). Error: " + exception, "PowerPoint LaTeX");
                fontFamily = new FontFamily("Times New Roman");
            }
            return fontFamily;
        }
    }
}
