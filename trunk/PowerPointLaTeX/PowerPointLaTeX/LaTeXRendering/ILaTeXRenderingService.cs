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
using System.Drawing;
using System.Windows.Forms;

namespace PowerPointLaTeX
{
    public struct LaTeXCompilationTask
    {
        /// <summary>
        /// (from the dvipng manpage)
        /// It reports the number of pixels from the bottom of the image to the baseline of the image.
        /// The depth is a negative offset in this case, so the minus sign is necessary, and the unit is pixels (px).
        /// </summary>
        public string code;

        /// <summary>
        /// target font size (in pixels)
        /// </summary>
        public float pixelsPerEmHeight;
    }

    public struct LaTeXCompilationResult
    {
        public byte[] imageData;

        public float baselineOffset;

        /// <summary>
        /// actual font size (in pixels)
        /// </summary>
        public float pixelsPerEmHeight;

        public string report;
    }

    public interface ILaTeXRenderingService
    {
        string AboutNotice {
            get;
        }

        string SeriveName {
            get;
        }

        /// <summary>
        /// Get the raw data of an image that can be read 
        /// </summary>
        /// <param name="latexCode"></param>
        /// <param name="image">the actual image of the rendered latexCode</param>
        /// <param name="baselineOffset"> 
        /// </param>
        /// <returns>returns false if there was an error</returns>
        LaTeXCompilationResult RenderLaTeXCode(LaTeXCompilationTask task);
    }
}
