using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System.Diagnostics;
using System.Windows.Forms;

/* Idea:
 * 
 * Create a lazy slide of a slide.
 * The lazy slide contains everything of the original slide (duplicate at first).
 * All changes made during editing to the original content as well as new content is stored in tags.
 * Either slide tags or in the big presentation tag.
 * As soon as the original slide is changed (anything in it), the lazy slide is adapted, too.
 * 
 * Prototype:
 *  Only text changes are allowed - only changes after the original text. Option to dim original text if new text is added.
 *  Option to replace the text entirely.
 *  
 */
namespace PowerPointLazySlides
{
    public partial class ThisAddIn
    {
        internal Presentation ActivePresentation
        {
            get
            {
                return Application.ActivePresentation;
            }
        }

        internal Slide ActiveSlide
        {
            get
            {
                // what about retrieving the current selection?
                // no, because multiple slides might have been selected but we are interested in the active one!
                return Application.ActiveWindow.View.Slide as Slide;
            }
        }

        private Dictionary<Presentation, Shape> lastTextShapeDict = new Dictionary<Presentation, Shape>();
        private Shape lastTextShape
        {
            get
            {
                if (!lastTextShapeDict.ContainsKey(ActivePresentation))
                {
                    return null;
                }
                return lastTextShapeDict[ActivePresentation];
            }
            set
            {
                lastTextShapeDict[ActivePresentation] = value;
            }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Application.WindowSelectionChange += new EApplication_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);
        }

        void Application_WindowSelectionChange(Selection Sel)
        {
            Shape currentShape = null;

            // we only care about text selections for now
            if (Sel.Type == PpSelectionType.ppSelectionText)
            {
                Debug.Assert(Sel.ShapeRange.Count == 1);
                currentShape = Sel.ShapeRange[1];
            }

            // grab the shape
            if (lastTextShape != currentShape)
            {
                if (lastTextShape != null)
                {
                    LeaveTextShape(lastTextShape);
                }
                if (currentShape != null)
                {
                    EnterTextShape(currentShape);
                }
            }
            lastTextShape = currentShape;
        }

        public void LeaveTextShape(Shape shape)
        {
            // make sure we have a parent slide
            Slide slide = null;
            try
            {
                slide = shape.Parent as Slide;
            }
            catch
            {
            }

            if (slide == null)
            {
                return;
            }

            // we only care about slides that are lazy slides here
            SlideTags tags = slide.LazySlideTags();
            ShapeTags shapeTags = shape.LazySlideTags();
            if (tags.IsLazySlide)
            {
                if (!shapeTags.ExcludeParentText)
                {
                    // remove the [] around the original text
                    TextRange inheritedRange = null;
                    // search for the first ] and replace that
                    TextRange range = shape.TextFrame.TextRange;
                    // TODO: this depends on GetMarkedText [1/5/2009 Andreas]
                    int startIndex = range.Text.IndexOf("[");
                    int endIndex = range.Text.IndexOf("]");
                    if (endIndex != -1 && startIndex == 0 && startIndex < endIndex)
                    {
                        inheritedRange = range.Characters(startIndex + 1, endIndex + 1);
                    }
                    else
                    {
                        // shit happens
                        MessageBox.Show("End of inherited text couldn't be found. Disabling inherited text for this shape.");
                        shapeTags.ExcludeParentText.value = true;

                    }

                    if (inheritedRange != null)
                    {
                        Slide parentSlide = FindSlideByLazySlideID(slide.LazySlideTags().ParentLazySlideId);
                        if (parentSlide == null)
                        {
                            return;
                        }

                        Shape parentShape = FindShapeByLazySlideID(parentSlide, shapeTags.ParentLazySlideId);
                        if (parentShape == null)
                        {
                            return;
                        }

                        TextRange parentRange = parentShape.TextFrame.TextRange;
                        CopyText(parentRange, inheritedRange);
                    }
                }
            }

            PropagateChanges(slide, shape);
        }

        public Shape GetCurrentTextShape()
        {
            Selection selection = Application.ActiveWindow.Selection;
            if (selection.Type != PpSelectionType.ppSelectionText)
            {
                return null;
            }
            return selection.ShapeRange[1];
        }

        public void IncludeInheritedText(Shape shape)
        {
            ShapeTags shapeTags = shape.LazySlideTags();

            Slide childSlide = shape.Parent as Slide;
            if (childSlide == null)
            {
                return;
            }

            Debug.Assert(childSlide.LazySlideTags().IsLazySlide);

            Slide parentSlide = FindSlideByLazySlideID(childSlide.LazySlideTags().ParentLazySlideId);
            if (parentSlide == null)
            {
                return;
            }

            Shape parentShape = FindShapeByLazySlideID(parentSlide, shapeTags.ParentLazySlideId);
            if (parentShape == null)
            {
                return;
            }

            if (shapeTags.ExcludeParentText)
            {
                // TODO: this code is duplicated from PropagateChanges.. [1/5/2009 Andreas]
                TextRange inheritedRange = parentShape.TextFrame.TextRange;
                TextRange childRange = shape.TextFrame.TextRange.Characters(1, 0);
                CopyText(inheritedRange, childRange);
                shapeTags.ParentTextLength.value = childRange.Length;

                shapeTags.ExcludeParentText.value = false;
            }
        }

        public void CopyText(TextRange parentRange, TextRange childRange)
        {
            childRange.Text = "";
            foreach (TextRange subRange in parentRange.Paragraphs(-1, -1))
            {
                TextRange childSubRange = childRange.InsertAfter(subRange.Text);
                childSubRange.IndentLevel = subRange.IndentLevel;
            }
        }

        public void ExcludeInheritedText(Shape shape)
        {
            ShapeTags shapeTags = shape.LazySlideTags();
            // not sure whether this is wanted behavior or not..
            /*
            if (!shapeTags.ExcludeParentText)
                        {
                            TextRange range = shape.TextFrame.TextRange;
                            if (GetCurrentTextShape() == shape)
                            {
                                TextRange range = shape.TextFrame.TextRange;
                                // TODO: this depends on GetMarkedText [1/5/2009 Andreas]
                                int startIndex = range.Text.IndexOf("[");
                                int endIndex = range.Text.IndexOf("]");
                                if (endIndex != -1 && startIndex != -1 && startIndex < endIndex)
                                {
                                    TextRange inheritedRange = range.Characters(startIndex + 1, endIndex + 1);
                                    inheritedRange.Text = "";
                                }
                            }
                            else
                            {
                                string inheritedText = shapeTags.ParentTextLength;
                                TextRange inheritedRange = range.Characters(1, inheritedText.Length);
                                inheritedRange.Text = "";
                            }
                        }*/


            shapeTags.ExcludeParentText.value = true;
        }

        private void PropagateChanges(Slide parentSlide, Shape parentShape)
        {
            if (!parentSlide.LazySlideTags().HasLazySlide)
            {
                return;
            }

            Slide childSlide = FindSlideByLazySlideID(parentSlide.LazySlideTags().ChildLazySlideId);
            if (childSlide == null)
            {
                return;
            }

            Shape childShape = FindShapeByLazySlideID(childSlide, parentShape.LazySlideTags().ChildLazySlideId);
            if (childShape == null)
            {
                return;
            }
            ShapeTags childShapeTags = childShape.LazySlideTags();
            if (childShapeTags.ExcludeParentText)
            {
                return;
            }

            TextRange parentRange = parentShape.TextFrame.TextRange;
            TextRange inheritedRange = childShape.TextFrame.TextRange.Characters(1, childShapeTags.ParentTextLength);
            CopyText(parentRange, inheritedRange);

            childShapeTags.ParentTextLength.value = inheritedRange.Length;

            PropagateChanges(childSlide, childShape);
        }

        public void EnterTextShape(Shape shape)
        {
            // make sure we have an active slide
            Slide slide = ActiveSlide;
            if (slide == null)
            {
                return;
            }

            // we only care about slides that are lazy slide
            SlideTags tags = slide.LazySlideTags();
            if (!tags.IsLazySlide)
            {
                return;
            }

            ShapeTags shapeTags = shape.LazySlideTags();
            Debug.Assert(shapeTags.LazySlideId != 0);
            if (!shapeTags.ExcludeParentText)
            {
                Slide parentSlide = FindSlideByLazySlideID(slide.LazySlideTags().ParentLazySlideId);
                if (parentSlide == null)
                {
                    return;
                }

                Shape parentShape = FindShapeByLazySlideID(parentSlide, shapeTags.ParentLazySlideId);
                if (parentShape == null)
                {
                    return;
                }

                TextRange parentRange = parentShape.TextFrame.TextRange;
                TextRange inheritedRange = shape.TextFrame.TextRange.Characters(1, shapeTags.ParentTextLength);
                CopyText(parentRange, inheritedRange);
                inheritedRange.InsertBefore("[");
                inheritedRange.InsertAfter("]");
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        public Slide FindSlideByLazySlideID(int id)
        {
            foreach (Slide slide in ActivePresentation.Slides)
            {
                if (slide.LazySlideTags().LazySlideId == id)
                {
                    return slide;
                }
            }
            return null;
        }

        public Shape FindShapeByLazySlideID(Slide slide, int id)
        {
            foreach (Shape shape in slide.Shapes)
            {
                if (shape.LazySlideTags().LazySlideId == id)
                {
                    return shape;
                }
            }
            return null;
        }

        public int GetNewLazySlideId()
        {
            PresentationTags tags = ActivePresentation.LazySlideTags();
            // is 0 at first and 0 is a reserved id (all non-lazy slides use it..)
            return ++tags.LazySlideIdCounter.value;
        }

        public void CreateLazySlideFor(Slide parentSlide)
        {
            Debug.Assert(parentSlide != null);

            SlideRange duplicateRange = parentSlide.Duplicate();
            Trace.Assert(duplicateRange.Count == 1);
            Slide childSlide = duplicateRange[1];

            // only set the tags after duplicating the slide..
            SlideTags parentTags = parentSlide.LazySlideTags();
            parentTags.HasLazySlide.value = true;
            if (parentTags.LazySlideId == 0)
            {
                parentTags.LazySlideId.value = GetNewLazySlideId();
            }

            SlideTags childTags = childSlide.LazySlideTags();
            childTags.IsLazySlide.value = true;
            childTags.LazySlideId.value = GetNewLazySlideId();

            childTags.ParentLazySlideId.value = parentTags.LazySlideId;
            parentTags.ChildLazySlideId.value = childTags.LazySlideId;

            // TODO: use the shape walker >_> [1/5/2009 Andreas]
            for (int i = 1; i <= parentSlide.Shapes.Count; i++)
            {
                Shape parentShape = parentSlide.Shapes[i];
                // only interested in text shapes
                if (parentShape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoFalse)
                {
                    continue;
                }

                Shape childShape = childSlide.Shapes[i];
                LinkLazyShapes(parentShape, childShape);
            }
        }

        private void LinkLazyShapes(Shape parentShape, Shape childShape)
        {
            ShapeTags parentTags = parentShape.LazySlideTags();
            if (parentTags.LazySlideId == 0)
            {
                parentTags.LazySlideId.value = GetNewLazySlideId();
            }

            // setup the text range length
            ShapeTags childTags = childShape.LazySlideTags();

            childTags.LazySlideId.value = GetNewLazySlideId();
            childTags.ParentLazySlideId.value = parentTags.LazySlideId;
            parentTags.ChildLazySlideId.value = childTags.LazySlideId;

            childTags.ParentTextLength.value = parentShape.TextFrame.TextRange.Length;
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
