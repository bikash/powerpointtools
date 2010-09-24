#region Copyright Notice
// This file is part of PowerPoint LaTeX.
// 
// Copyright (C) 2008/2009 Andreas Kirsch
// 
// PowerPoint LaTeX is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
// 
// PowerPoint LaTeX is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
// 
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <http://www.gnu.org/licenses/>.
#endregion

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;
using System.IO;
using System.Drawing;
using System.Diagnostics;
using System.Windows.Forms;
using System.Threading;

namespace PowerPointLaTeX
{
    public enum EquationType
    {
        None,
        // has inline LaTeX codes (but not compiled)
        HasInlines,
        // has compiled LaTeX codes
        HasCompiledInlines,
        // a compiled LaTeX code element (picture)
        Inline,
        // an equation (picture)
        Equation,
    }

    static class EquationTypeShapeExtension
    {
        public static bool IsEquation(this Shape shape)
        {
            return shape.LaTeXTags().Type == EquationType.Equation;
        }
    }

    /// <summary>
    /// Contains all the important methods, etc.
    /// Instantiated by the add-in
    /// </summary>
    class LaTeXTool : PowerPointLaTeX.ILaTeXTool
    {
        private const int InitialEquationFontSize = 44;

        private Microsoft.Office.Interop.PowerPoint.Application Application
        {
            get
            {
                return Globals.ThisAddIn.Application;
            }
        }

        internal Presentation ActivePresentation
        {
            get { return Application.ActivePresentation; }
        }

        internal Slide ActiveSlide
        {
            get { return Application.ActiveWindow.View.Slide as Slide; }
        }

        /// <summary>
        /// returns whether the addin is enabled in the current context (ie presentation)
        /// but is also affected by the global addin settings, of course.
        /// </summary>
        internal bool AddInEnabled
        {
            get
            {
                return !ActivePresentation.Final && Properties.Settings.Default.EnableAddIn && Compatibility.IsSupportedPresentation( ActivePresentation );
            }
        }

        /// <summary>
        /// Compile latexCode into an inline shape
        /// </summary>
        /// <param name="slide"></param>
        /// <param name="textShape"></param>
        /// <param name="latexCode"></param>
        /// <param name="codeRange"></param>
        /// <returns></returns>
        private Shape CompileInlineLaTeXCode(Slide slide, Shape textShape, string latexCode, TextRange codeRange)
        {
            Shape picture = LaTeXRendering.GetPictureShapeFromLaTeXCode( slide, latexCode, codeRange.Font.Size );
            if (picture == null)
            {
                return null;
            }

            // add tags to the picture
            picture.LaTeXTags().Code.value = latexCode;
            picture.LaTeXTags().Type.value = EquationType.Inline;
            picture.LaTeXTags().LinkID.value = textShape.Id;

            InlineFormulas.Alignment.PrepareTextRange(picture, codeRange);

            CopyInlineEffects( slide, textShape, codeRange, picture );

            return picture;
        }

        private void CopyInlineEffects( Slide slide, Shape textShape, TextRange codeRange, Shape picture ) {
            try {
                // copy animations from the parent textShape
                Sequence sequence = slide.TimeLine.MainSequence;
                var effects =
                    from Effect effect in sequence
                    where effect.Shape.SafeThis() != null && effect.Shape == textShape &&
                    ((effect.EffectInformation.TextUnitEffect == MsoAnimTextUnitEffect.msoAnimTextUnitEffectByParagraph &&
                        textShape.ParagraphContainsRange( effect.GetSafeParagraph(), codeRange ))
                        || effect.EffectInformation.BuildByLevelEffect == MsoAnimateByLevel.msoAnimateLevelNone)
                    select effect;

                picture.AddEffects(effects, true, sequence);
            }
            catch {
                Debug.Fail( "CopyInlineEffects failed!" );
            }
        }



        private void CompileInlineTextRange(Slide slide, Shape shape, TextRange range)
        {
            int startIndex = 0;

            int codeCount = 0;

            List<TextRange> pictureRanges = new List<TextRange>();
            List<Shape> pictures = new List<Shape>();

            while (true)
            {
                bool inlineMode;
                int latexCodeStartIndex;
                int latexCodeEndIndex;
                int endIndex;

                int inlineStartIndex, displaystyleStartIndex;
                inlineStartIndex = range.Text.IndexOf("$$", startIndex);
                displaystyleStartIndex = range.Text.IndexOf( "$$[", startIndex );
                if( displaystyleStartIndex != -1 && displaystyleStartIndex <= inlineStartIndex ) {
                    inlineMode = false;

                    startIndex = displaystyleStartIndex;
                    latexCodeStartIndex = startIndex + 3;
                    latexCodeEndIndex = range.Text.IndexOf( "]$$", latexCodeStartIndex );
                    endIndex = latexCodeEndIndex + 3;
                } else if( inlineStartIndex != -1 ) {
                    inlineMode = true;

                    startIndex = inlineStartIndex;
                    latexCodeStartIndex = startIndex + 2;
                    latexCodeEndIndex = range.Text.IndexOf( "$$", latexCodeStartIndex );
                    endIndex = latexCodeEndIndex + 2;
                }
                else {
                    break;
                }

                if( latexCodeEndIndex == -1 )
                {
                    break;
                }

                int length = endIndex - startIndex;

                int latexCodeLength = latexCodeEndIndex - latexCodeStartIndex;
                string latexCode = range.Text.Substring( latexCodeStartIndex, latexCodeLength );
                latexCode = TransformSpecialUnicodeCharactersToAscii(latexCode);

                // must be [[ then
                if( !inlineMode ) {
                    latexCode = @"\displaystyle{" + latexCode + "}";
                }

                LaTeXEntry tagEntry = shape.LaTeXTags().Entries[codeCount];
                tagEntry.Code.value = latexCode;
                // TODO: cohesion? [5/2/2009 Andreas]
                // save the font size because it might be changed later
                // +1 because IndexOf is base 0, but Characters uses base 1
                tagEntry.FontSize.value = range.Characters( latexCodeStartIndex + 1, latexCodeLength ).Font.Size;

                // escape $$!$$
                // +1 because IndexOf is base 0, but Characters uses base 1
                TextRange codeRange = range.Characters(startIndex + 1, length);
                if (!Helpers.IsEscapeCode(latexCode))
                {
                    Shape picture = CompileInlineLaTeXCode(slide, shape, latexCode, codeRange);
                    if (picture != null)
                    {
                        tagEntry.ShapeId.value = picture.Id;

                        pictures.Add(picture);
                        pictureRanges.Add(codeRange);
                    }
                    else
                    {
                        codeRange.Text = "$Formula Error$";
                    }
                }
                else
                {
                    codeRange.Text = "$$";
                }

                tagEntry.StartIndex.value = codeRange.Start;
                tagEntry.Length.value = codeRange.Length;

                // NOTE: endIndex isnt valid anymore since we've removed some text [5/24/2010 Andreas
                // IndexOf uses base0, codeRange base1 => -1
                startIndex = codeRange.Start + codeRange.Length - 1;
                codeCount++;
            }

            // now that everything has been converted we can position the formulas (pictures) in the text area
            for (int i = 0; i < pictures.Count; i++)
            {
                TextRange codeRange = pictureRanges[i];
                InlineFormulas.Alignment.Align(codeRange, pictures[i]);
            }

            // update the type, too
            shape.LaTeXTags().Type.value = codeCount > 0 ? EquationType.HasCompiledInlines : EquationType.None;
        }

        private string TransformSpecialUnicodeCharactersToAscii(string text)
        {
            text = text.Replace((char)8217, '\'');
            // replace weird unicode - (hypens) with minus
            text = text.Replace((char)8208, '-');
            text = text.Replace((char)8211, '-');
            text = text.Replace((char)8212, '-');
            text = text.Replace((char)8722, '-');
            text = text.Replace((char)8209, '-');
            text = text.Replace((char)8259, '-');
            // replace ellipses with ...
            text = text.Replace(((char)8230).ToString(), "...");

            return text;
        }

        private void DecompileTextRange(Slide slide, Shape shape, TextRange range)
        {
            // make sure this is always valid, otherwise the code will do stupid things
            Debug.Assert(shape.LaTeXTags().Type == EquationType.HasCompiledInlines);

            LaTeXEntries entries = shape.LaTeXTags().Entries;
            int length = entries.Length;
            for (int i = length - 1; i >= 0; i--)
            {
                LaTeXEntry entry = entries[i];
                string latexCode = entry.Code;

                if (!Helpers.IsEscapeCode(latexCode))
                {
                    int shapeID = entry.ShapeId;
                    // find the shape
                    Shape picture = slide.Shapes.FindById(shapeID);

                    //Debug.Assert(picture != null);
                    // fail gracefully
                    if (picture != null)
                    {
                        Debug.Assert(picture.LaTeXTags().Type == EquationType.Inline);
                        picture.Delete();
                    }

                    // release the cache entry, too
                    ActivePresentation.CacheTags()[latexCode].Release();
                }

                // add back the latex code
                TextRange codeRange = range.Characters(entry.StartIndex, entry.Length);
                if( latexCode.StartsWith( @"\displaystyle{" ) && latexCode.EndsWith( "}" ) ) {
                    codeRange.Text = "$$[" + latexCode.Substring( @"\displayStyle{".Length, latexCode.Length - 1 - @"\displayStyle{".Length ) + "]$$";
                }
                else {
                    codeRange.Text = "$$" + latexCode + "$$";
                }
                if (entry.FontSize != 0) {
                    codeRange.Font.Size = entry.FontSize;
                }
                codeRange.Font.BaselineOffset = 0.0f;
            }

            entries.Clear();
            shape.LaTeXTags().Type.value = EquationType.HasInlines;
        }

        public void CompileShape(Slide slide, Shape shape)
        {
            // we don't need to compile already compiled shapes (its also sensible to avoid destroying escape sequences or overwrite entries, etc)
            // don't try to compile equations (or their sources) either
            EquationType type = shape.LaTeXTags().Type;
            if (type == EquationType.HasCompiledInlines || type == EquationType.Equation)
            {
                return;
            }

            if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                TextFrame textFrame = shape.TextFrame;
                if (textFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    CompileInlineTextRange(slide, shape, textFrame.TextRange);
                }
            }
        }

        public void DecompileShape(Slide slide, Shape shape)
        {
            // we don't need to decompile already shapes that aren't compiled
            if (shape.LaTeXTags().Type != EquationType.HasCompiledInlines)
            {
                return;
            }

            if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                TextFrame textFrame = shape.TextFrame;
                if (textFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    DecompileTextRange(slide, shape, textFrame.TextRange);
                }
            }
        }

        private void FinalizeShape(Slide slide, Shape shape)
        {
            CompileShape(slide, shape);
            shape.LaTeXTags().Clear();
        }

        public void CompileSlide(Slide slide)
        {
            ShapeWalker.WalkSlide(slide, CompileShape);
        }

        public void DecompileSlide(Slide slide)
        {
            ShapeWalker.WalkSlide(slide, DecompileShape);
        }

        public void CompilePresentation(Presentation presentation)
        {
            ShapeWalker.WalkPresentation(presentation, CompileShape);
        }

        /// <summary>
        /// Removes all tags and all pictures that belong to inline formulas
        /// </summary>
        /// <param name="slide"></param>
        public void FinalizePresentation(Presentation presentation)
        {
            ShapeWalker.WalkPresentation(presentation, FinalizeShape);
            // purge the cache, too
            presentation.CacheTags().PurgeAll();
            presentation.SettingsTags().Clear();
        }

        public Shape CreateEmptyEquation()
        {
            const float width = 100, height = 60;

            Shape shape = ActiveSlide.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 100, 100, width, height);
            shape.Fill.ForeColor.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorBackground1;
            shape.Fill.Solid();

            shape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;

            LaTeXTags tags = shape.LaTeXTags();
            tags.Type.value = EquationType.Equation;
            tags.OriginalWidth.value = width;
            tags.OriginalHeight.value = height;
            tags.FontSize.value = InitialEquationFontSize;

            return shape;
        }

        public Shape EditEquation( Shape equation, out bool cancelled ) {
            EquationEditor editor = new EquationEditor( equation.LaTeXTags().Code, equation.LaTeXTags().FontSize );
            DialogResult result = editor.ShowDialog();
            if( result == DialogResult.Cancel ) {
                cancelled = true;
                // don't change anything
                return equation;
            }
            else {
                cancelled = false;
            }

            // recompile the code
            //equation.TextFrame.TextRange.Text = equationSource.TextFrame.TextRange.Text;
            string latexCode = editor.LaTeXCode;

            Slide slide = equation.GetSlide();
            if (slide == null) {
                // TODO: what do we do in this case? [3/3/2009 Andreas]
                return equation;
            }

            Shape newEquation = null;
            if (latexCode.Trim() != "") {
                newEquation = LaTeXRendering.GetPictureShapeFromLaTeXCode( slide, latexCode, editor.FontSize );
            }

            if (newEquation != null) {
                LaTeXTags tags = newEquation.LaTeXTags();

                tags.OriginalWidth.value = newEquation.Width;
                tags.OriginalHeight.value = newEquation.Height;
                tags.FontSize.value = editor.FontSize;

                tags.Type.value = EquationType.Equation;
            }
            else {
                newEquation = CreateEmptyEquation();
            }

            newEquation.LaTeXTags().Code.value = latexCode;

            newEquation.Top = equation.Top;
            newEquation.Left = equation.Left;

            // keep the equation's scale
            // TODO: this scales everything twice if we are not careful [3/4/2009 Andreas]
            float widthScale = equation.Width / equation.LaTeXTags().OriginalWidth;
            float heightScale = equation.Height / equation.LaTeXTags().OriginalHeight;
            newEquation.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
            newEquation.Width *= widthScale;
            newEquation.Height *= heightScale;

            // copy animations over from the old equation
            Sequence sequence = slide.TimeLine.MainSequence;
            var effects =
                from Effect effect in sequence
                where effect.Shape == equation
                select effect;

            newEquation.AddEffects(effects, false, sequence);

            // delete the old equation
            equation.Delete();

            // return the new equation
            return newEquation;
        }
    }
}
