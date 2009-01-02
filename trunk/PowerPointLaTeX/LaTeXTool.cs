﻿using System;
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
    /// <summary>
    /// Contains all the important methods, etc.
    /// Instantiated by the add-in
    /// </summary>
    class LaTeXTool
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

        private Shape CompileLaTeXCode(Slide slide, Shape shape, string latexCode, TextRange codeRange)
        {
            // check the cache first
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
            Shape picture = AddPictureFromData(slide, imageData);

            picture.AlternativeText = latexCode;

            //pictureRange.Width = range.BoundWidth;
            //pictureRange.Height = range.BoundHeight;

            // add tags to the picture
            picture.LaTeXTags().Code.value = latexCode;
            picture.LaTeXTags().Type.value = EquationType.Inline;
            picture.LaTeXTags().ParentId.value = shape.Id;

            // align the picture and remove the original text
            // 1 Point = 1/72 Inches
            float fontSize = codeRange.Font.Size;
            // TODO: erm... [12/30/2008 Andreas]
            /*
            float scalingFactor = fontSize / 50.0f;
                        picture.Width *= scalingFactor;
                        picture.Height *= scalingFactor;*/


            /*
              System.Drawing.Font font = new System.Drawing.Font(codeRange.Font.Name, fontSize, GraphicsUnit.Point);
                         FontFamily fontFamily = new FontFamily(codeRange.Font.Name);
                         fontFamily.*/

            // base line: center (assume that its a one-line codeRange)
            if (fontSize > picture.Height)
            {
                picture.Top = codeRange.BoundTop + (fontSize - picture.Height);

            }
            else
            {
                picture.Top = codeRange.BoundTop + (fontSize - picture.Height) * 0.5f;
            }

            // fill up text with spaces to "wrap around" the object
            codeRange.Text = "";
            float width = 0;
            while (width < picture.Width)
            {
                codeRange.Text += " ";
                width = codeRange.BoundWidth;
            }

            picture.Left = codeRange.BoundLeft;

            // copy animations from the parent shape
            // TODO: braindead API [12/31/2008 Andreas]

            return picture;
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
            if (codeCount > 0)
            {
                shape.LaTeXTags().Type.value = EquationType.HasCompiledInlines;
            }
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
            shape.LaTeXTags().Clear();
        }

        public void DecompileShape(Slide slide, Shape shape)
        {
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
