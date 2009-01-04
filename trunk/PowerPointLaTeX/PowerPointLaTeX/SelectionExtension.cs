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
        static internal void SelectShapes<T>(this Selection selection, T list, bool replace) where T : IEnumerable<Shape>
        {
            if (replace)
            {
                selection.Unselect();
            }

            foreach (Shape shape in list)
            {
                shape.Select(Microsoft.Office.Core.MsoTriState.msoFalse);
            }
        }

        /// <summary>
        /// Get a shapes of all selected shapes (included ones that are selected indirectly through text selections)
        /// Slides don't select all shapes though.
        /// </summary>
        /// <param name="selection"></param>
        /// <returns></returns>
        static internal List<Shape> GetShapesFromShapeSelection(this Selection selection)
        {
            List<Shape> shapes = new List<Shape>();
            if (selection.Type == PpSelectionType.ppSelectionShapes)
            {
                foreach (Shape shape in selection.ShapeRange)
                {
                    shapes.Add(shape);
                }
            }

            return shapes;
        }

        static internal Shape GetShapeFromTextSelection(this Selection selection)
        {
            Shape shape = null;
            if (selection.Type == PpSelectionType.ppSelectionText)
            {
                Trace.Assert(selection.ShapeRange.Count == 1);
                shape = selection.ShapeRange[1];
            }
            return shape;
        }

        static internal void FilterShapes(this Selection selection, System.Predicate<Shape> predicate)
        {
            if (selection.Type != PpSelectionType.ppSelectionShapes)
            {
                return;
            }

            List<Shape> shapes = selection.GetShapesFromShapeSelection();
            shapes = shapes.FindAll(predicate);
            selection.SelectShapes(shapes, false);
        }
    }
}
