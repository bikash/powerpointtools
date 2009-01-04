using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;
using System.IO;
using System.Drawing;
using System.Diagnostics;
using System.Windows.Forms;

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
        Equation
    }

    static class EquationTypeExtension
    {
        /// <summary>
        /// Returns whether the type hints that the shape contains addin-specific data
        /// </summary>
        /// <returns></returns>
        internal static bool ContainsLaTeXCode(this EquationType type)
        {
            return type != EquationType.None;
        }
    }

    /// <summary>
    /// Contains all the important methods, etc.
    /// Instantiated by the add-in
    /// </summary>
    class LaTeXTool
    {
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

        internal bool EnableAddIn {
            get
            {
                return !ActivePresentation.Final && Properties.Settings.Default.EnableAddIn;
            }
        }

        // TODO: rename the stupid webservice faff! [12/30/2008 Andreas]
        private Shape AddPictureFromData(Slide slide, byte[] data)
        {
            MemoryStream stream = new MemoryStream(data);
            Image image = Image.FromStream(stream);

            IDataObject oldClipboardContent = Clipboard.GetDataObject();

            Clipboard.Clear();
            Clipboard.SetImage(image);

            ShapeRange pictureRange = slide.Shapes.Paste();
            Clipboard.SetDataObject(oldClipboardContent);

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


        private Shape CompileLaTeXCode(Slide slide, Shape shape, string latexCode, TextRange codeRange)
        {
            // check the cache first
            byte[] imageData;
            imageData = GetImageForLaTeXCode(latexCode);
            Shape picture = AddPictureFromData(slide, imageData);

            picture.AlternativeText = latexCode;
            picture.Name = "LaTeX: " + latexCode;

            // add tags to the picture
            picture.LaTeXTags().Code.value = latexCode;
            picture.LaTeXTags().Type.value = EquationType.Inline;
            picture.LaTeXTags().ParentId.value = shape.Id;

            FitFormulaIntoText(codeRange, picture);

            // copy animations from the parent shape
            Sequence sequence = slide.TimeLine.MainSequence;
            var allEffect = from Effect e in sequence select e;
            var effects =
                from Effect effect in sequence
                where effect.Shape == shape &&
                (effect.EffectInformation.TextUnitEffect == MsoAnimTextUnitEffect.msoAnimTextUnitEffectByParagraph &&
                    ParagraphContainsRange(shape, effect.Paragraph, codeRange))
                select effect;

            foreach (Effect effect in effects)
            {
                int index = effect.Index + 1;
                Effect formulaEffect = sequence.Clone(effect, index);
                formulaEffect = sequence.ConvertToBuildLevel(formulaEffect, MsoAnimateByLevel.msoAnimateLevelNone);
                //formulaEffect = sequence.ConvertToTextUnitEffect(formulaEffect, MsoAnimTextUnitEffect.msoAnimTextUnitEffectMixed);
                formulaEffect.Timing.TriggerType = MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                formulaEffect.Paragraph = 0;
                formulaEffect.Shape = picture;
                // Effect formulaEffect = sequence.AddEffect(picture, effect.EffectType, MsoAnimateByLevel.msoAnimateLevelNone, MsoAnimTriggerType.msoAnimTriggerWithPrevious, index);
            }

            return picture;
        }

        private static void FitFormulaIntoText(TextRange codeRange, Shape picture)
        {
            // interesting fact: text filled with spaces -> BoundHeight == EmSize
            codeRange.Text = " ";
            float fontHeight = codeRange.BoundHeight;
            FontFamily fontFamily = new FontFamily(codeRange.Font.Name);
            float baselineHeight = (float) (fontHeight * ((float) fontFamily.GetCellAscent(FontStyle.Regular) / fontFamily.GetLineSpacing(FontStyle.Regular)));

            float scalingFactor = codeRange.Font.Size / 44.0f; // baselineHeight / 34.5f;
            // NOTE: PowerPoint scales both Width and Height if you scale one of them >_< [1/2/2009 Andreas]
            picture.Height *= scalingFactor;

            // fill the text up with spaces to make it "wrap around" the formula
            while (codeRange.BoundWidth < picture.Width)
            {
                codeRange.Text += " ";
            }

            // align the picture
            picture.Left = codeRange.BoundLeft;
            if (baselineHeight >= picture.Height)
            {
                // baseline: center (assume that its a one-line codeRange)
                picture.Top = codeRange.BoundTop + baselineHeight - picture.Height;
            }
            else
            {
                picture.Top = codeRange.BoundTop + (fontHeight - picture.Height) * 0.5f;
            }
        }

        private byte[] GetImageForLaTeXCode(string latexCode)
        {
            byte[] imageData;
            if (ActivePresentation.CacheTags()[latexCode].IsCached())
            {
                imageData = ActivePresentation.CacheTags()[latexCode].Use();
            }
            else
            {
                LaTeXWebService.WebService.URLData URLData = LaTeXWebService.WebService.compileLaTeX(latexCode);
                Trace.Assert(System.Text.RegularExpressions.Regex.IsMatch(URLData.contentType, "gif|bmp|jpeg|png"));
                imageData = URLData.content;

                ActivePresentation.CacheTags()[latexCode].Store(imageData);
            }
            return imageData;
        }

        private bool IsEscapeCode(string code)
        {
            return code == "!";
        }

        private void CompileTextRange(Slide slide, Shape shape, TextRange range)
        {
            int startIndex = 0;

            int codeCount = 0;
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

                LaTeXEntry tagEntry = shape.LaTeXTags().Entries[codeCount];
                tagEntry.Code.value = latexCode;

                // escape $$!$$
                TextRange codeRange = range.Characters(startIndex + 1 - 2, length + 4);
                if (!IsEscapeCode(latexCode))
                {
                    Shape picture = CompileLaTeXCode(slide, shape, latexCode, codeRange);
                    tagEntry.ShapeId.value = picture.Id;
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

            // update the type, too
            shape.LaTeXTags().Type.value = codeCount > 0 ? EquationType.HasCompiledInlines : EquationType.None;
        }

        public void CompileShape(Slide slide, Shape shape)
        {
            // we don't need to compile already compiled shapes (its also sensible to avoid destroying escape sequences or overwrite entries, etc.)
            if (shape.LaTeXTags().Type == EquationType.HasCompiledInlines)
            {
                return;
            }

            if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                TextFrame textFrame = shape.TextFrame;
                if (textFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    CompileTextRange(slide, shape, textFrame.TextRange);
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

                    Debug.Assert(picture != null);
                    Debug.Assert(picture.LaTeXTags().Type == EquationType.Inline);
                    // fail gracefully
                    if (picture != null)
                    {
                        picture.Delete();
                    }

                    // release the cache entry, too
                    ActivePresentation.CacheTags()[latexCode].Release();
                }

                // add back the latex code
                TextRange codeRange = range.Characters(entry.StartIndex, entry.Length);
                codeRange.Text = "$$" + latexCode + "$$";
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
                    Trace.Assert(inlineShape != null);
                    shapes.Add(inlineShape);
                }
            }
            return shapes;
        }

        /// <summary>
        /// from an inline shape
        /// </summary>
        /// <param name="shape"></param>
        /// <returns></returns>
        public Shape GetParentShape(Shape shape)
        {
            LaTeXTags tags = shape.LaTeXTags();
            Debug.Assert(tags.Type == EquationType.Inline);
            Slide slide = shape.GetSlide();
            Trace.Assert(slide != null);

            Shape parent = slide.Shapes.FindById(tags.ParentId);
            Trace.Assert(parent != null);
            return parent;
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
