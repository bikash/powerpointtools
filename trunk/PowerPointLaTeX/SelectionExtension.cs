using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;
using System.Diagnostics;

namespace PowerPointLaTeX
{
    static class SelectionExtension
    {
        static internal void SelectShapes<T>(this Selection selection, T list) where T : IEnumerable<Shape>
        {
            selection.Unselect();

            foreach (Shape shape in list)
            {
                shape.Select(Microsoft.Office.Core.MsoTriState.msoFalse);
            }
        }

        /// <summary>
        /// Get a list of all selected shapes (included ones that are selected indirectly through text selections)
        /// Slides don't select all shapes though.
        /// </summary>
        /// <param name="selection"></param>
        /// <returns></returns>
        static internal List<Shape> GetShapes(this Selection selection)
        {
            List<Shape> shapes = new List<Shape>();
            switch (selection.Type)
            {
                case PpSelectionType.ppSelectionShapes:
                    foreach (Shape shape in selection.ShapeRange)
                    {
                        shapes.Add(shape);
                    }
                    break;
                case PpSelectionType.ppSelectionText:
                    Trace.Assert(selection.ShapeRange.Count == 1);
                    shapes.Add(selection.ShapeRange[1]);
                    break;
                default:
                    break;
            }

            return shapes;
        }

        static internal void FilterShapes(this Selection selection, System.Predicate<Shape> predicate)
        {
            if (selection.Type != PpSelectionType.ppSelectionShapes)
            {
                return;
            }

            List<Shape> shapes = selection.GetShapes();
            shapes = shapes.FindAll(predicate);
            selection.SelectShapes(shapes);
        }
    }
}
