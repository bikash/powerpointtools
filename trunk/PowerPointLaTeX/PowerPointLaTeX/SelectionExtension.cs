#region Copyright Notice
// This file is part of PowerPoint LaTeX.
// 
// Copyright (C) 2008/2009 Andreas Kirsch
// 
// PowerPoint LaTeX is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
// 
// PowerPoint LaTeX is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
// 
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <http://www.gnu.org/licenses/>.
#endregion

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
        /// Get a shapes of all selected shapes (excluding the shape selected through text selections)
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

        /// <summary>
        /// Get the shapes from the current selection.
        /// Text and shape selections are straight-forward
        /// A slide selection returns all shapes of the selected slides.
        /// </summary>
        /// <param name="selection"></param>
        /// <returns></returns>
        static internal List<Shape> GetShapes(this Selection selection) {
            List<Shape> shapes;
            if (selection.Type == PpSelectionType.ppSelectionShapes)
            {
                shapes = selection.GetShapesFromShapeSelection();
            }
            else
            {
                shapes = new List<Shape>();
                if (selection.Type == PpSelectionType.ppSelectionText)
                {
                    shapes.Add(selection.GetShapeFromTextSelection());
                }
                else if (selection.Type == PpSelectionType.ppSelectionSlides)
                {
                    foreach( Slide slide in selection.SlideRange ) {
                        foreach( Shape shape in slide.Shapes) {
                            shapes.Add(shape);
                        }
                    }
                }
            }
            return shapes;
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
