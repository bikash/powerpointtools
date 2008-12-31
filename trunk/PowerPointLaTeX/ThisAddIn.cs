using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;
using System.Diagnostics;

namespace PowerPointLaTeX
{
    public partial class ThisAddIn
    {
        internal LaTeXTool Tool
        {
            get;
            private set;
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Tool = new LaTeXTool();

            // register events
            Application.PresentationSave += new EApplication_PresentationSaveEventHandler(Application_PresentationSave);
            Application.SlideShowBegin += new EApplication_SlideShowBeginEventHandler(Application_SlideShowBegin);
            Application.WindowBeforeDoubleClick += new EApplication_WindowBeforeDoubleClickEventHandler(Application_WindowBeforeDoubleClick);
            Application.WindowSelectionChange += new EApplication_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);
        }

        private Shape oldTextShape = null;

        void Application_WindowSelectionChange(Selection Sel)
        {
            // automatically select the parent (and thus all children) of a inline objects
            List<Shape> shapes = Sel.GetShapesFromShapeSelection();
            // all parent shapes of inline shapes in shapes
            IEnumerable<Shape> parentShapes =
                shapes.FindAll(shape => shape.LaTeXTags().Type == LaTeXTool.EquationType.Inline).
                    ConvertAll(inlineShape => Tool.GetParentShape(inlineShape));
            // add the parent shapes in the original selection
            parentShapes = parentShapes.Union(
                    shapes.FindAll(shape => shape.LaTeXTags().Type == LaTeXTool.EquationType.HasCompiledInlines)
                );

            IEnumerable<Shape> shapeUnion = shapes.Union(parentShapes);

            foreach (Shape shape in parentShapes)
            {
                List<Shape> inlineShapes = Tool.GetInlineShapes(shape);
                shapeUnion = shapeUnion.Union(inlineShapes);
            }
            Sel.SelectShapes(shapeUnion, false);

            Shape textShape = Sel.GetShapeFromTextSelection();
            // recompile the old shape if necessary (do nothing if we click around in the same text shape though)
            if( oldTextShape != null && oldTextShape != textShape) {
                Tool.CompileShape(oldTextShape.GetSlide(), oldTextShape);
            }
            if (textShape != null)
            {
                Tool.DecompileShape(Tool.ActiveSlide, textShape);
            }

            oldTextShape = textShape;
        }

        void Application_WindowBeforeDoubleClick(Selection Sel, ref bool Cancel)
        {
            // kill the MS engineers - kill them all and torture them slowly to death..
            // http://www.eggheadcafe.com/software/aspnet/33533167/ppt--windowbeforedoublec.aspx
        }

        void Application_SlideShowBegin(SlideShowWindow Wn)
        {
            //throw new NotImplementedException();
        }

        void Application_PresentationSave(Presentation Pres)
        {
            //throw new NotImplementedException();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }


        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
