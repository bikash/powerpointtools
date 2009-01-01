using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Windows.Forms.Design;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLaTeX
{
    public partial class DeveloperTaskPaneControl : UserControl
    {
        public static string Title {
            get { return "LaTeX Developer Pane"; }
        }

        public DeveloperTaskPaneControl()
        {
            InitializeComponent();

            Globals.ThisAddIn.Application.WindowSelectionChange += new EApplication_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);
        }

        void Application_WindowSelectionChange(Selection Sel)
        {
            ControlCollection controls = tagsLayout.Controls;

            controls.Clear();
            switch (Sel.Type)
            {
                case PpSelectionType.ppSelectionShapes:
                case PpSelectionType.ppSelectionText:
                    foreach (Shape shape in Sel.ShapeRange)
                    {
                        controls.Add(new TagsGrid(String.Format("Shape {0}:", shape.Id), shape.Tags));
                    }
                    break;
                case PpSelectionType.ppSelectionSlides:
                    foreach (Slide slide in Sel.SlideRange){
                        controls.Add(new TagsGrid(String.Format("Shape {0} ({1}):", slide.Shapes.Title, slide.SlideID), slide.Tags));
                    }
                    break;
                case PpSelectionType.ppSelectionNone:
                    Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
                    controls.Add(new TagsGrid(String.Format("Presentation {0}", presentation.Name), presentation.Tags));
                    break;
            }
        }
    }
}
