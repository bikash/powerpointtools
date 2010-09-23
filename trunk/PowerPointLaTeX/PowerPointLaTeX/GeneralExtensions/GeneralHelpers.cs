using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLaTeX
{
    static class GeneralHelpers
    {
        static public bool RangesOverlap(TextRange rangeA, TextRange rangeB)
        {
            int startA = rangeA.Start;
            int endA = startA + rangeA.Length - 1;
            int startB = rangeB.Start;
            int endB = startB + rangeB.Length - 1;
            return !(endA < startB || endB < startA);
        }

        static public bool ParagraphContainsRange(Shape shape, int paragraph, TextRange range)
        {
            TextRange paragraphRange = shape.TextFrame.TextRange.Paragraphs(paragraph, 1);
            return RangesOverlap(paragraphRange, range);
        }
    }
}
