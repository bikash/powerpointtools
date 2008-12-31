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
        internal static Slide GetSlide(this Shape shape) {
            Slide parent = shape.Parent as Slide;
            Trace.Assert(parent != null);
            return parent;
        }
    }
}
