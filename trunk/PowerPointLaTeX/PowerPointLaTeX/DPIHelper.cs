using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PowerPointLaTeX
{
    // TODO: find a better name for the class? [9/22/2010 Andreas]
    class DPIHelper
    {
        public const int WindowsDPISetting = 96;
        public const int PrintPtsPerInch = 72;

        public static float FontSizeToPixelsPerEmHeight(float fontSizeInPoints, int targetPtsPerInch)
        {
            return fontSizeInPoints / PrintPtsPerInch * targetPtsPerInch;
        }

        public static float FontSizeToPixelsPerEmHeight(float fontSizeInPoints)
        {
            return FontSizeToPixelsPerEmHeight(fontSizeInPoints, WindowsDPISetting);
        }

        public static float PixelsPerEmHeightToFontSize(float pixelHeight, int sourcePtsPerInch)
        {
            return pixelHeight / sourcePtsPerInch * PrintPtsPerInch;
        }

        public static float PixelsPerEmHeightToFontSize(float pixelHeight)
        {
            return PixelsPerEmHeightToFontSize( pixelHeight, WindowsDPISetting );
        }
    }
}
