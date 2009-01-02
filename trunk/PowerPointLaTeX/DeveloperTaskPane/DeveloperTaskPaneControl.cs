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
        public static string Title
        {
            get { return "LaTeX Developer Pane"; }
        }

        private Presentation Presentation
        {
            get
            {
                return Globals.ThisAddIn.Application.ActivePresentation;
            }
        }

        public DeveloperTaskPaneControl()
        {
            InitializeComponent();

            domPropertyGrid.PropertyTabs.AddTabType(typeof(ReflectionPropertyTab), PropertyTabScope.Global);

            Globals.ThisAddIn.Application.WindowSelectionChange += new EApplication_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);
        }

        class ReflectionPropertyTab : PropertyTab {
            public override PropertyDescriptorCollection GetProperties(object component, Attribute[] attributes)
            {
                return ReflectionPropertyDescriptor.FromObject(component, attributes);
            }

            public override string TabName
            {
                get { return "Fields"; }
            }

            public override Bitmap Bitmap
            {
                get
                {
                    return new System.Drawing.Bitmap(16, 16);
                }
            }
        }

        class SelectionWrapper {
            public Selection selection {
                get;
                private set;
            }

            public SelectionWrapper(Selection selection) {
                this.selection = selection;
            }
        }

        void Application_WindowSelectionChange(Selection Sel)
        {
            domPropertyGrid.SelectedObject = new SelectionWrapper( Sel );

            ProcessSelection(Sel);
        }

        private void ProcessSelection(Selection Sel)
        {
            tagsLayout.Controls.Clear();
            // an exception is thrown otherwise >_>
            if (Presentation.Final)
            {
                return;
            }
            switch (Sel.Type)
            {
                case PpSelectionType.ppSelectionShapes:
                case PpSelectionType.ppSelectionText:
                    AddShapesToSelection(Sel.ShapeRange);
                    break;
                case PpSelectionType.ppSelectionSlides:
                    AddSlidesToSelection(Sel.SlideRange);
                    break;
                case PpSelectionType.ppSelectionNone:
                    AddPresentationToSelection();
                    break;
            }
        }

        private void AddShapesToSelection(System.Collections.IEnumerable shapes)
        {
            foreach (Shape shape in shapes)
            {
                tagsLayout.Controls.Add(new TagsGrid(String.Format("Shape {0}:", shape.Id), shape.Tags));
            }
        }

        private void AddSlidesToSelection(System.Collections.IEnumerable slides)
        {
            foreach (Slide slide in slides)
            {
                tagsLayout.Controls.Add(new TagsGrid(String.Format("Shape {0} ({1}):", slide.Shapes.Title, slide.SlideID), slide.Tags));
            }
        }

        private void AddPresentationToSelection()
        {
            tagsLayout.Controls.Add(new TagsGrid(String.Format("Presentation {0}", Presentation.Name), Presentation.Tags));
        }

        private void refreshButton_Click(object sender, EventArgs e)
        {
            foreach(TagsGrid tagsGrid in tagsLayout.Controls) {
                tagsGrid.RefreshTags();
            }
        }

        private void selectAllButton_Click(object sender, EventArgs e)
        {
            tagsLayout.Controls.Clear();

            AddPresentationToSelection();
            AddSlidesToSelection(Presentation.Slides);
            foreach (Slide slide in Presentation.Slides)
            {
                AddShapesToSelection(slide.Shapes);
            }
        }

        private void useCurrentSelectionButton_Click(object sender, EventArgs e)
        {
            ProcessSelection(Globals.ThisAddIn.Application.ActiveWindow.Selection);
        }

    }
}
