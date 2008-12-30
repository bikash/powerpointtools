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
        private const string TagPrefix = "PowerPointLaTeX_";
        private const string CodeTag = "Code";
        private const string TypeTag = "Type";

        // used for equations that will be toggled
        private const string TypeInline = "inline";
        // used for equations that will be kept (separate object that is edited with the equation editor)
        private const string TypeEquation = "equation";

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

        private void CompileLaTeXCode(Slide slide, string latexCode, TextRange codeRange)
        {
            LaTeXWebService.WebService.URLData data = LaTeXWebService.WebService.compileLaTeX(latexCode);
            Shape picture = AddPictureFromData(slide, data);

            picture.AlternativeText = latexCode;

            //pictureRange.Width = range.BoundWidth;
            //pictureRange.Height = range.BoundHeight;

            picture.Tags.Add(TagPrefix + CodeTag, latexCode);
            picture.Tags.Add(TagPrefix + TypeTag, TypeInline);
            
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
        }

        private void CompileTextRange(Slide slide, TextRange range)
        {
            string text = range.Text;
            int startIndex = 0;
            while ((startIndex = text.IndexOf("$$", startIndex)) != -1)
            {
                startIndex += 2;

                int endIndex = text.IndexOf("$$", startIndex);
                if (endIndex == -1)
                {
                    break;
                }

                int length = endIndex - startIndex;
                string latexCode = text.Substring(startIndex, length);

                TextRange codeRange = range.Characters(startIndex + 1 - 2, length + 4);
                CompileLaTeXCode(slide, latexCode, codeRange);

                startIndex = endIndex + 2;
            }
        }

        private void CompileShape(Slide slide, Shape shape)
        {
            if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                TextFrame textFrame = shape.TextFrame;
                if (textFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    CompileTextRange(slide, textFrame.TextRange);
                }
            }
            else if (shape.HasTable == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                Table table = shape.Table;
                foreach (Row row in table.Rows)
                {
                    foreach (Cell cell in row.Cells)
                    {
                        CompileShape(slide, cell.Shape);
                    }
                }
            }
        }

        public void CompileSlide(Slide slide)
        {
            foreach (Shape shape in slide.Shapes)
            {
                CompileShape(slide, shape);
            }
        }
    }
}
