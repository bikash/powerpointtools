using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.PowerPoint;
using System.Diagnostics;
using System.Windows.Forms;
using System.IO;
using System.Drawing;

namespace PowerPointLaTeX
{
    public partial class LaTeXRibbon : OfficeRibbon
    {
        public LaTeXRibbon()
        {
            InitializeComponent();
        }

        private void LaTeXRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Slide slide = ((Slide)Globals.ThisAddIn.Application.ActiveWindow.View.Slide);
            foreach (Shape shape in slide.Shapes)
            {
                if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoFalse)
                {
                    continue;
                }

                TextFrame textFrame = shape.TextFrame;
                if (textFrame.HasText == Microsoft.Office.Core.MsoTriState.msoFalse)
                {
                    continue;
                }

                string text = textFrame.TextRange.Text;
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
                    //MessageBox.Show(latexCode);

                    TextRange range = textFrame.TextRange.Characters(startIndex + 1 - 2, length + 4);
                    LaTeXWebService.WebService.URLData data = LaTeXWebService.WebService.compileLaTeX(@"\displaystyle{ " + latexCode + "}");
                    Trace.Assert(data.contentType.Contains("png"));

                    MemoryStream stream = new MemoryStream(data.content);
                    Image image = Image.FromStream(stream);

                    IDataObject oldClipboardContent = Clipboard.GetDataObject();

                    Clipboard.Clear();
                    Clipboard.SetImage(image);

                    ShapeRange pictureRange = slide.Shapes.Paste();
                    pictureRange.AlternativeText = latexCode;

                    //pictureRange.Width = range.BoundWidth;
                    //pictureRange.Height = range.BoundHeight;

                    pictureRange.Tags.Add("WebLaTeXCode", latexCode);
                    
                    //range.Font.Size = range.BoundWidth / pictureRange
                    pictureRange.Left = range.BoundLeft;
                    pictureRange.Top = range.BoundTop;

                    Clipboard.SetDataObject(oldClipboardContent);
                    startIndex = endIndex + 2;
                }
            }

            /*
                       string latexCode = @"a \le b";
                        string hexData = "";
                               foreach (byte c in data.content) {
                                   hexData += c.ToString("X2");
                               }
                               //{\*\shppict }
                               Clipboard.SetText(@"{\pict \pngblip " + hexData + "}", TextDataFormat.Rtf);
                               //range.Paste();
                               TextRange picture = range.PasteSpecial(PpPasteDataType.ppPasteRTF,Microsoft.Office.Core.MsoTriState.msoFalse,"",0,"",Microsoft.Office.Core.MsoTriState.msoFalse);
 
                        */

        }
    }
}
