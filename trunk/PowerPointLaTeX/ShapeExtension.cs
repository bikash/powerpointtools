using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;
using System.Diagnostics;

namespace PowerPointLaTeX
{
    static class ShapeExtension
    {
        /// <summary>
        /// Might return null!
        /// </summary>
        /// <param name="shape"></param>
        /// <returns></returns>
        internal static Slide GetSlide(this Shape shape) {
            Slide parent = shape.Parent as Slide;
            return parent;
        }
    }
}
