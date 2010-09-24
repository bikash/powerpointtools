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

namespace PowerPointLaTeX
{
    class ShapeWalker
    {
        public delegate void ShapeVisitor(Slide slide, Shape shape);

        static public void WalkShape(Slide slide, Shape shape, ShapeVisitor doShape)
        {
            if (shape.HasTable == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                Table table = shape.Table;
                foreach (Row row in table.Rows)
                {
                    foreach (Cell cell in row.Cells)
                    {
                        WalkShape(slide, cell.Shape, doShape);
                    }
                }
            }
            // TODO: group chapes and childshaperanges are not supported yet! [5/25/2010 Andreas]
            /*
            foreach( Shape subShape in shape.GroupItems ) {
                WalkShape( slide, subShape, doShape );
            }*/


            doShape(slide, shape);
        }

        static public void WalkSlide(Slide slide, ShapeVisitor walkTextRange)
        {
            foreach (Shape shape in slide.Shapes)
            {
                WalkShape(slide, shape, walkTextRange);
            }
        }

        static public void WalkPresentation(Presentation presentation, ShapeVisitor walkTextRange)
        {
            foreach (Slide slide in presentation.Slides)
            {
                WalkSlide(slide, walkTextRange);
            }
        }
    }
}
