using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLaTeX
{
    static class ShapesExtension
    {
        internal static Shape FindById(this Shapes shapes, int shapeID)
        {
            Shape result = null;

            foreach (Shape shape in shapes)
            {
                if (shape.Id == shapeID)
                {
                    result = shape;
                    break;
                }
            }

            return result;
        }
    }
}
