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
    /// <summary>
    /// Contains all the important methods, etc.
    /// Instantiated by the add-in
    /// </summary>
    class LaTeXTool
    {
        public enum EquationType
        {
            None,
            HasInlines,
            HasCompiledInlines,
            Inline,
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

        // TODO: rename the stupid webservice faff! [12/30/2008 Andreas]
        private Shape AddPictureFromData(Slide slide, LaTeXWebService.WebService.URLData data)
        {
            Trace.Assert(System.Text.RegularExpressions.Regex.IsMatch(data.contentType, "gif|bmp|jpeg|png"));

            MemoryStream stream = new MemoryStream(data.content);
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
            LaTeXWebService.WebService.URLData data = LaTeXWebService.WebService.compileLaTeX(latexCode);
            Shape picture = AddPictureFromData(slide, data);

            picture.AlternativeText = latexCode;

            //pictureRange.Width = range.BoundWidth;
            //pictureRange.Height = range.BoundHeight;

            // add tags to the picture
            picture.LaTeXTags().Code = latexCode;
            picture.LaTeXTags().Type = EquationType.Inline;

            // align the picture and remove the original text
            // 1 Point = 1/72 Inches
            float fontSize = codeRange.Font.Size;
            // TODO: erm... [12/30/2008 Andreas]
            /*
            float scalingFactor = fontSize / 50.0f;
                        picture.Width *= scalingFactor;
                        picture.Height *= scalingFactor;*/


            //System.Drawing.Font font = new System.Drawing.Font(codeRange.Font.Name, fontSize, GraphicsUnit.Point);

            // base line: center (assume that its a one-line codeRange)
            picture.Top = codeRange.BoundTop + (fontSize - picture.Height) * 0.5f;

            codeRange.ParagraphFormat.LineRuleWithin = Microsoft.Office.Core.MsoTriState.msoFalse;
            codeRange.ParagraphFormat.SpaceWithin = picture.Height;

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
                tagEntry.Code = latexCode;

                // escape $$!$$
                TextRange codeRange = range.Characters(startIndex + 1 - 2, length + 4);
                if (latexCode != "!")
                {
                    Shape picture = CompileLaTeXCode(slide, shape, latexCode, codeRange);
                    tagEntry.ShapeID = picture.Id;
                }
                else
                {
                    codeRange.Text = "$$";
                }

                tagEntry.StartIndex = codeRange.Start;
                tagEntry.Length = codeRange.Length;

                startIndex = codeRange.Start + codeRange.Length - 1;
                codeCount++;
            }
            if (codeCount > 0)
            {
                shape.LaTeXTags().Type = EquationType.HasCompiledInlines;
            }
        }

        private void CompileShape(Slide slide, Shape shape)
        {
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

                if (latexCode != "!")
                {
                    int shapeID = entry.ShapeID;
                    // find the shape
                    Shape picture = slide.Shapes.FindById(shapeID);

                    Debug.Assert(picture != null);
                    Debug.Assert(picture.LaTeXTags().Type == EquationType.Inline);
                    // fail gracefully
                    if (picture != null)
                    {
                        picture.Delete();
                    }
                }

                // add back the latex code
                TextRange codeRange = range.Characters(entry.StartIndex, entry.Length);
                codeRange.Text = "$$" + latexCode + "$$";
            }
            shape.LaTeXTags().Clear();
        }

        private void DecompileShape(Slide slide, Shape shape)
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

        private void BakeShape(Slide slide, Shape shape)
        {
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

        public void CompileSlide(Slide slide)
        {
            WalkSlide(slide, CompileShape);
        }

        public void DecompileSlide(Slide slide)
        {
            WalkSlide(slide, DecompileShape);
        }

        /// <summary>
        /// Removes all tags and all pictures that belong to inline formulas
        /// </summary>
        /// <param name="slide"></param>
        public void BakeSlide(Slide slide)
        {
            WalkSlide(slide, BakeShape);
        }
    }
}
