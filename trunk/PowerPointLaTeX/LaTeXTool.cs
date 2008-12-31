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
        public enum EquationType {
            None,
            Inline,
            Equation
        }

        private delegate void WalkTextRange(Slide slide, Shape shape, TextRange textRange);
               
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
            codeRange.Text = "";
            float width = 0;
            while (width < picture.Width)
            {
                codeRange.Text += " ";
                width = codeRange.BoundWidth;
            }

            picture.Left = codeRange.BoundLeft;

            return picture;
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
                
                TextRange codeRange = range.Characters(startIndex + 1 - 2, length + 4);
                Shape picture = CompileLaTeXCode(slide, shape, latexCode, codeRange);

                tagEntry.StartIndex = codeRange.Start;
                tagEntry.Length = codeRange.Length;
                tagEntry.ShapeID = picture.Id;
                
                startIndex = codeRange.Start + codeRange.Length - 1;
                codeCount++;
            }
        }        

        private void DecompileTextRange(Slide slide, Shape shape, TextRange range) {
            LaTeXEntries entries = shape.LaTeXTags().Entries;
            int length = entries.Length;
            for( int i = length - 1 ; i >= 0 ; i-- ) {
                LaTeXEntry entry = entries[i];
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

                // add back the latex code
                TextRange codeRange = range.Characters(entry.StartIndex, entry.Length);
                codeRange.Text = "$$" + entry.Code + "$$";
            }
            shape.LaTeXTags().Clear();
        }

        private void PurgeInlinesFromTextRange(Slide slide, Shape shape, TextRange range)
        {
            if (shape.LaTeXTags().Type == EquationType.None)
            {
                shape.LaTeXTags().Clear();
            }
        }

        private void WalkShape(Slide slide, Shape shape, WalkTextRange walkTextRange)
        {
            if (shape.HasTable == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                Table table = shape.Table;
                foreach (Row row in table.Rows)
                {
                    foreach (Cell cell in row.Cells)
                    {
                        WalkShape(slide, cell.Shape, walkTextRange);
                    }
                }
            }
            if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                TextFrame textFrame = shape.TextFrame;
                if (textFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    walkTextRange(slide, shape, textFrame.TextRange);
                }
            }
        }

        private void WalkSlide(Slide slide, WalkTextRange walkTextRange)
        {
            foreach (Shape shape in slide.Shapes)
            {
                WalkShape(slide, shape, walkTextRange);
            }
        }

        public void CompileSlide(Slide slide)
        {
            WalkSlide(slide, CompileTextRange);
        }

        public void DecompileSlide(Slide slide) {
            WalkSlide(slide, DecompileTextRange);
        }

        /// <summary>
        /// Removes all tags and all pictures that belong to inline formulas
        /// </summary>
        /// <param name="slide"></param>
        public void PurgeInlinesFromSlide(Slide slide) {
            WalkSlide(slide, PurgeInlinesFromTextRange);

            // purge all inline pictures, too
            foreach(Shape shape in slide.Shapes) {
                if( shape.LaTeXTags().Type == EquationType.Inline ) {
                    shape.Delete();
                }
            }
        }
    }
}
