using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;
using System.Diagnostics;

namespace PowerPointLaTeX
{
    static class ShapeExtensions
    {
        internal static List<Shape> GetInlineShapes(Shape shape)
        {
            List<Shape> shapes = new List<Shape>();

            Slide slide = shape.GetSlide();
            if (slide == null)
            {
                return shapes;
            }

            foreach (LaTeXEntry entry in shape.LaTeXTags().Entries)
            {
                if (!IsEscapeCode(entry.Code))
                {
                    Shape inlineShape = slide.Shapes.FindById(entry.ShapeId);
                    Debug.Assert(inlineShape != null);
                    if (inlineShape != null)
                    {
                        shapes.Add(inlineShape);
                    }
                }
            }
            return shapes;
        }

        /// <summary>
        /// get the shape specified by the LinkID field in LaTeXTags()
        /// </summary>
        /// <param name="shape"></param>
        /// <returns></returns>
        internal static Shape GetLinkShape(this Shape shape)
        {
            LaTeXTags tags = shape.LaTeXTags();
            // pre-condition: shape has a linked shape
            Debug.Assert(tags.Type == EquationType.Inline || tags.Type == EquationType.Equation);
            Slide slide = shape.GetSlide();
            Trace.Assert(slide != null);

            Shape linkShape = slide.Shapes.FindById(tags.LinkID);
            Trace.Assert(linkShape != null);
            return linkShape;
        }

    }
}
