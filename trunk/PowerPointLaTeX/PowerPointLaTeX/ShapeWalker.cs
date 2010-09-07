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
