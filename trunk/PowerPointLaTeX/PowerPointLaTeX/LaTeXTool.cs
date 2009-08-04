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
        // a compiled LaTeX code (picture)
        Inline,
        // an equation (not implemented atm)
        Equation,
    }

    static class EquationTypeShapeExtension
    {
        internal static bool IsEquation(this Shape shape)
        {
            return shape.LaTeXTags().Type == EquationType.Equation;
        }
    }

    /// <summary>
    /// Contains all the important methods, etc.
    /// Instantiated by the add-in
    /// </summary>
    class LaTeXTool
    {
        private const char NoneBreakingSpace = (char) 160;

        private delegate void DoShape(Slide slide, Shape shape);

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

        internal bool EnableAddIn
        {
            get
            {
                return !ActivePresentation.Final && Properties.Settings.Default.EnableAddIn;
            }
        }

        /// <returns>A valid picture shape from the data or null if creation failed</returns>
        private Shape GetPictureShapeFromImage(Slide slide, Image image)
        {           
            IDataObject oldClipboardContent = null;
            try {
                oldClipboardContent = Clipboard.GetDataObject();
            }
            catch { Debug.Assert( false, "Retrieving the current clipboard contents failed!"); }

            Clipboard.SetImage(image);

            ShapeRange pictureRange = slide.Shapes.Paste();
            if( oldClipboardContent != null )
                Clipboard.SetDataObject(oldClipboardContent);

            if( pictureRange == null ) {
                return null;
            }

            // make white the transparent color
            pictureRange.PictureFormat.TransparencyColor = ~0;
            pictureRange.PictureFormat.TransparentBackground = Microsoft.Office.Core.MsoTriState.msoCTrue;

            Trace.Assert(pictureRange.Count == 1);
            return pictureRange[1];
        }

        private bool RangesOverlap(TextRange rangeA, TextRange rangeB)
        {
            int startA = rangeA.Start;
            int endA = startA + rangeA.Length - 1;
            int startB = rangeB.Start;
            int endB = startB + rangeB.Length - 1;
            return !(endA < startB || endB < startA);
        }

        private bool ParagraphContainsRange(Shape shape, int paragraph, TextRange range)
        {
            TextRange paragraphRange = shape.TextFrame.TextRange.Paragraphs(paragraph, 1);
            return RangesOverlap(paragraphRange, range);
        }

        private Shape GetPictureShapeFromLaTeXCode(Slide currentSlide, string latexCode)
        {
            // check the cache first
            Image image = GetImageForLaTeXCode(latexCode);
            if (image == null) {
                return null;
            }

            Shape picture = GetPictureShapeFromImage(currentSlide, image);
            if( picture == null ) {
                return null;
            }

            picture.AlternativeText = latexCode;

            string shortenedName;
            if (latexCode.Length > 32) {
                shortenedName = latexCode.Substring(0, 32) + "..";
            }
            else {
                shortenedName = latexCode;
            }
            picture.Name = "LaTeX: " + shortenedName;

            return picture;
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
            Shape picture = GetPictureShapeFromLaTeXCode(slide, latexCode);
            if (picture == null)
            {
                return null;
            }

            // add tags to the picture
            picture.LaTeXTags().Code.value = latexCode;
            picture.LaTeXTags().Type.value = EquationType.Inline;
            picture.LaTeXTags().LinkID.value = textShape.Id;
            
            // scale the picture to fit the font size
            // TODO: magic numbers? [4/17/2009 Andreas]
            float scalingFactor = codeRange.Font.Size / 44.0f; // baselineHeight / 34.5f;
            picture.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
            picture.Height *= scalingFactor;
            picture.Width *= scalingFactor;

            // change the font size to keep the formula from overlapping with regular text (nifty :))
            if (codeRange.Font.Size < picture.Height) {
                codeRange.Font.Size = picture.Height;
            }
            codeRange.Font.BaselineOffset = -0.5f;

            // disable word wrap
            codeRange.ParagraphFormat.WordWrap = Microsoft.Office.Core.MsoTriState.msoFalse;
            // fill the text up with none breaking space to make it "wrap around" the formula
            FillTextRange(codeRange, NoneBreakingSpace, picture.Width);

            // copy animations from the parent textShape
            Sequence sequence = slide.TimeLine.MainSequence;
            var effects =
                from Effect effect in sequence
                where effect.Shape == textShape &&
                ((effect.EffectInformation.TextUnitEffect == MsoAnimTextUnitEffect.msoAnimTextUnitEffectByParagraph &&
                    ParagraphContainsRange(textShape, effect.Paragraph, codeRange))
                    || effect.EffectInformation.BuildByLevelEffect == MsoAnimateByLevel.msoAnimateLevelNone)
                select effect;

            CopyEffectsTo(picture, true, sequence, effects);

            return picture;
        }

        private static void CopyEffectsTo(Shape target, bool setToWithPrevious, Sequence sequence, IEnumerable<Effect> effects)
        {
            foreach (Effect effect in effects)
            {
                int index = effect.Index + 1;
                Effect formulaEffect = sequence.Clone(effect, index);
                try
                {
                    formulaEffect = sequence.ConvertToBuildLevel(formulaEffect, MsoAnimateByLevel.msoAnimateLevelNone);
                }
                catch { }
                //formulaEffect = sequence.ConvertToTextUnitEffect(formulaEffect, MsoAnimTextUnitEffect.msoAnimTextUnitEffectMixed);
                if (setToWithPrevious)
                    formulaEffect.Timing.TriggerType = MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                try
                {
                    formulaEffect.Paragraph = 0;
                }
                catch { }
                formulaEffect.Shape = target;
                // Effect formulaEffect = sequence.AddEffect(picture, effect.EffectType, MsoAnimateByLevel.msoAnimateLevelNone, MsoAnimTriggerType.msoAnimTriggerWithPrevious, index);
            }
        }

        private static void FillTextRange(TextRange range, char character, float minWidth)
        {
            range.Text = character.ToString();

            // line-breaks are futile, so break if one happen 
            float oldHeight = range.BoundHeight;
            while (range.BoundWidth < minWidth && oldHeight == range.BoundHeight)
            {
                range.Text += character.ToString();
            }
            if( oldHeight != range.BoundHeight ) {
                range.Text = range.Text.Remove(0, 1);
            }
        }

        private static void AlignFormulaWithText(TextRange codeRange, Shape picture)
        {
            // interesting fact: text filled with (at most one line of none-breaking) spaces -> BoundHeight == EmSize
            //codeRange.Text = " ";
            float fontHeight = codeRange.BoundHeight;
            FontFamily fontFamily = new FontFamily(codeRange.Font.Name);
            float baselineHeight = (float) (fontHeight * ((float) fontFamily.GetCellAscent(FontStyle.Regular) / fontFamily.GetLineSpacing(FontStyle.Regular)));

            picture.Left = codeRange.BoundLeft;

            if (baselineHeight >= picture.Height)
            {
                // baseline: center (assume that its a one-line codeRange)
                picture.Top = codeRange.BoundTop + (baselineHeight - picture.Height) * 0.5f;
            }
            else
            {
                // center the picture directly
                picture.Top = codeRange.BoundTop + (fontHeight - picture.Height) * 0.5f;
            }
        }

        /// <summary>
        /// Get the raw image data for some latex code or null if the compilation failed.
        /// It seems to be possible to receive a byte[0] array for some reason.
        /// </summary>
        /// <param name="latexCode"></param>
        /// <returns></returns>
        private byte[] GetImageDataForLaTeXCode(string latexCode)
        {
            byte[] imageData;
            // TODO: rewrite the cache system to work even if the main thread is blocked [8/4/2009 Andreas]
            if (ActivePresentation.CacheTags()[latexCode].IsCached())
            {
                imageData = ActivePresentation.CacheTags()[latexCode].Use();
            }
            else
            {
                imageData = Globals.ThisAddIn.LaTeXServices.Service.GetImageDataForLaTeXCode(latexCode);
                if (imageData == null)
                {
                    return null;
                }

                ActivePresentation.CacheTags()[latexCode].Store(imageData);
            }

            // make sure we return a some-what meaningful array
            Debug.Assert(imageData.Length > 0);
            return imageData;
        }


        private Image GetImageFromImageData(byte[] imageData) {
            if( imageData == null ) {
                return null;
            }

            MemoryStream stream = new MemoryStream(imageData);
            Image image;
            try {
                image = Image.FromStream(stream);
            }
            catch {
                return null;
            }
            return image;
        }

        public Image GetImageForLaTeXCode(string latexCode) {
            byte[] imageData = GetImageDataForLaTeXCode(latexCode);
            return GetImageFromImageData(imageData);
        }

        private static bool IsEscapeCode(string code)
        {
            return code == "!";
        }

        private void CompileInlineTextRange(Slide slide, Shape shape, TextRange range)
        {
            int startIndex = 0;

            int codeCount = 0;

            List<TextRange> pictureRanges = new List<TextRange>();
            List<Shape> pictures = new List<Shape>();

            while ((startIndex = range.Text.IndexOf("$$", startIndex)) != -1)
            {
                startIndex += 2;

                int endIndex = range.Text.IndexOf("$$", startIndex);
                if (endIndex == -1)
                {
                    break;
                }

                int length = endIndex - startIndex;
                string latexCode = range.Text.Substring(startIndex, length);
                // TODO: move this into its own function [1/5/2009 Andreas]
                // replace weird unicode ' with the one usually used
                latexCode = latexCode.Replace((char) 8217, '\'');

                LaTeXEntry tagEntry = shape.LaTeXTags().Entries[codeCount];
                tagEntry.Code.value = latexCode;
                // TODO: cohesion? [5/2/2009 Andreas]
                // save the font size because it might be changed later
                tagEntry.FontSize.value = range.Characters(startIndex, length).Font.Size;

                // escape $$!$$
                TextRange codeRange = range.Characters(startIndex + 1 - 2, length + 4);
                if (!IsEscapeCode(latexCode))
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

                startIndex = codeRange.Start + codeRange.Length - 1;
                codeCount++;
            }

            // TODO: this doesn't work - simply disable autofit instead.. [1/5/2009 Andreas]
            // TODO: can we automate this? [2/26/2009 Andreas]
            /*
                   (new Thread( delegate() {
                            Thread.Sleep(100);
                            for (int i = 0; i < pictures.Count; i++)
                            {
                                TextRange codeRange = pictureRanges[i];
                                AlignFormulaWithText(codeRange, pictures[i]);
                            }
                        } )).Start();*/

            // now that everything has been converted we can position the formulas (pictures) in the text area
            for (int i = 0; i < pictures.Count; i++)
            {
                TextRange codeRange = pictureRanges[i];
                AlignFormulaWithText(codeRange, pictures[i]);
            }

            // update the type, too
            shape.LaTeXTags().Type.value = codeCount > 0 ? EquationType.HasCompiledInlines : EquationType.None;
        }

        public void CompileShape(Slide slide, Shape shape)
        {
            // we don't need to compile already compiled shapes (its also sensible to avoid destroying escape sequences or overwrite entries, etc.)
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

                if (!IsEscapeCode(latexCode))
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
                    ActivePresentation.CacheTags()[latexCode].ReleaseIfUsed();
                }

                // add back the latex code
                TextRange codeRange = range.Characters(entry.StartIndex, entry.Length);
                codeRange.Text = "$$" + latexCode + "$$";
                if (entry.FontSize != 0) {
                    codeRange.Font.Size = entry.FontSize;
                }
                codeRange.Font.BaselineOffset = 0.0f;
            }

            entries.Clear();
            shape.LaTeXTags().Type.value = EquationType.HasInlines;
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

        /*
          private bool presentationNeedsCompile;
                 // TODO: use an exception, etc. to early-out [1/2/2009 Andreas]
                 private void ShapeNeedsCompileWalker(Slide slide, Shape shape) {
                     EquationType type = shape.LaTeXTags().Type;
                     if( type == EquationType.HasInlines ) {
                         presentationNeedsCompile = true;
                     }
                     if( type == EquationType.None ) {
                         // check it and assign a type if necessary
                         if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                         {
                             TextFrame textFrame = shape.TextFrame;
                             if (textFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                             {
                                 // search for a $$ - if there is one occurrence, it needs to be compiled
                                 if(textFrame.TextRange.Text.IndexOf("$$") != -1) {
                                     presentationNeedsCompile = true;
                                     // update the type, too
                                     shape.LaTeXTags().Type.value = EquationType.HasInlines;
                                 } else {
                                     shape.LaTeXTags().Type.value = EquationType.None;
                                 }
                             }
                         }
                     }
                 }
        
                 public bool NeedsCompile(Presentation presentation) {
                     presentationNeedsCompile = false;
                     WalkPresentation(presentation, ShapeNeedsCompileWalker);
                     return presentationNeedsCompile;
                 }*/


        private void FinalizeShape(Slide slide, Shape shape)
        {
            CompileShape(slide, shape);
            shape.LaTeXTags().Clear();
        }

        private void WalkShape(Slide slide, Shape shape, DoShape doShape)
        {
            if (shape.HasTable == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                Table table = shape.Table;
                foreach (Row row in table.Rows)
                {
                    foreach (Cell cell in row.Cells)
                    {
                        WalkShape(slide, cell.Shape, doShape);
                    }
                }
            }
            doShape(slide, shape);
        }

        private void WalkSlide(Slide slide, DoShape walkTextRange)
        {
            foreach (Shape shape in slide.Shapes)
            {
                WalkShape(slide, shape, walkTextRange);
            }
        }

        private void WalkPresentation(Presentation presentation, DoShape walkTextRange)
        {
            foreach (Slide slide in presentation.Slides)
            {
                WalkSlide(slide, walkTextRange);
            }
        }

        public void CompileSlide(Slide slide)
        {
            WalkSlide(slide, CompileShape);
        }

        public void DecompileSlide(Slide slide)
        {
            WalkSlide(slide, DecompileShape);
        }

        public void CompilePresentation(Presentation presentation)
        {
            WalkPresentation(presentation, CompileShape);
        }

        /// <summary>
        /// Removes all tags and all pictures that belong to inline formulas
        /// </summary>
        /// <param name="slide"></param>
        public void FinalizePresentation(Presentation presentation)
        {
            WalkPresentation(presentation, FinalizeShape);
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

            return shape;
        }

        public Shape EditEquation( Shape equation ) {
            EquationEditor editor = new EquationEditor(equation.LaTeXTags().Code);
            DialogResult result = editor.ShowDialog();
            if( result == DialogResult.Cancel ) {
                // don't change anything
                return equation;
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
                newEquation = GetPictureShapeFromLaTeXCode(slide, latexCode);
            }

            if (newEquation != null) {
                LaTeXTags tags = newEquation.LaTeXTags();

                tags.OriginalWidth.value = newEquation.Width;
                tags.OriginalHeight.value = newEquation.Height;
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

            CopyEffectsTo(newEquation, false, sequence, effects);

            // delete the old equation
            equation.Delete();

            // return the new equation
            return newEquation;
        }

        public List<Shape> GetInlineShapes(Shape shape)
        {
            List<Shape> shapes = new List<Shape>();

            Slide slide = shape.GetSlide();
            if (slide == null)
            {
                return shapes;
            }

            foreach (LaTeXEntry entry in shape.LaTeXTags().Entries)
            {
                if (!IsEscapeCode(entry.Code))
                {
                    Shape inlineShape = slide.Shapes.FindById(entry.ShapeId);
                    Debug.Assert(inlineShape != null);
                    if (inlineShape != null)
                    {
                        shapes.Add(inlineShape);
                    }
                }
            }
            return shapes;
        }

        /// <summary>
        /// get the shape specified by the LinkID field in LaTeXTags()
        /// </summary>
        /// <param name="shape"></param>
        /// <returns></returns>
        public Shape GetLinkShape(Shape shape)
        {
            LaTeXTags tags = shape.LaTeXTags();
            Debug.Assert(tags.Type == EquationType.Inline || tags.Type == EquationType.Equation);
            Slide slide = shape.GetSlide();
            Trace.Assert(slide != null);

            Shape linkShape = slide.Shapes.FindById(tags.LinkID);
            Trace.Assert(linkShape != null);
            return linkShape;
        }

        private ShapeRange GetShapeRange<T>(T list) where T : IEnumerable<Shape>
        {
            Selection selection = Application.ActiveWindow.Selection;
            selection.Unselect();

            foreach (Shape shape in list)
            {
                shape.Select(Microsoft.Office.Core.MsoTriState.msoFalse);
            }

            // FIXME: ignore ChildShapeRange for now [12/31/2008 Andreas]
            ShapeRange shapeRange = selection.ShapeRange;
            selection.Unselect();

            return shapeRange;
        }
    }
}
