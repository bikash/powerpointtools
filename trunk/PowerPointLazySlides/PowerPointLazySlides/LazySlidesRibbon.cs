using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLazySlides
{
    public partial class LazySlidesRibbon : OfficeRibbon
    {
        private ThisAddIn AddIn
        {
            get { return Globals.ThisAddIn; }
        }
        public LazySlidesRibbon()
        {
            InitializeComponent();
        }

        private void LazySlidesRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            AddIn.Application.WindowSelectionChange += new EApplication_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);
        }

        void Application_WindowSelectionChange(Selection Sel)
        {
            Selection selection = AddIn.Application.ActiveWindow.Selection;

            if (selection.Type != PpSelectionType.ppSelectionText &&
                selection.Type != PpSelectionType.ppSelectionShapes ||
                selection.ShapeRange.Count != 1)
            {
                IncludeInheritedTextToggle.Enabled = false;
                IncludeInheritedTextToggle.Checked = false;
            }

            Shape shape = selection.ShapeRange[1];
            // TODO: duplicate code.. [1/5/2009 Andreas]
            if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoFalse)
            {
                IncludeInheritedTextToggle.Enabled = false;
                IncludeInheritedTextToggle.Checked = false;
                return;
            }

            Slide childSlide = shape.Parent as Slide;
            if (childSlide == null || !childSlide.LazySlideTags().IsLazySlide)
            {
                IncludeInheritedTextToggle.Enabled = false;
                IncludeInheritedTextToggle.Checked = false;
                return;
            }

            IncludeInheritedTextToggle.Enabled = true;
            IncludeInheritedTextToggle.Checked = !shape.LazySlideTags().ExcludeParentText;
        }

        private void createLazySlideButton_Click(object sender, RibbonControlEventArgs e)
        {
            Slide slide = AddIn.ActiveSlide;
            if (slide == null)
            {
                return;
            }

            SlideTags tags = slide.LazySlideTags();
            if (tags.HasLazySlide)
            {
                // TODO: check if the child slide has been deleted by accident.. [1/5/2009 Andreas]
                return;
            }

            AddIn.CreateLazySlideFor(slide);
        }

        private void IncludeInheritedTextToggle_Click(object sender, RibbonControlEventArgs e)
        {
            // TODO: add a ActiveSelection property! [1/5/2009 Andreas]
            Selection selection = AddIn.Application.ActiveWindow.Selection;

            if (selection.Type != PpSelectionType.ppSelectionText &&
                selection.Type != PpSelectionType.ppSelectionShapes ||
                selection.ShapeRange.Count != 1)
            {
                return;
            }
            
            Shape shape = selection.ShapeRange[1];

            if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoFalse)
            {
                return;
            }

            Slide childSlide = shape.Parent as Slide;
            if (childSlide == null || !childSlide.LazySlideTags().IsLazySlide)
            {
                return;
            }

            // TODO: hack? see the selec in the end [1/5/2009 Andreas]
            selection.Unselect();

            if (IncludeInheritedTextToggle.Checked)
            {
                AddIn.IncludeInheritedText(shape);
            }
            else
            {
                AddIn.ExcludeInheritedText(shape);
            }
        }
    }
}
