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
