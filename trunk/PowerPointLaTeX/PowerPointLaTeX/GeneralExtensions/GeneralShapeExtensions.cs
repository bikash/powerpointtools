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
    static class GeneralShapeExtensions
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

        /// <summary>
        /// Make sure that the shape still exists and return null otherwise.
        /// </summary>
        /// <param name="shape"></param>
        /// <returns>Return the shape itself or null if it doesn't exist anymore</returns>
        internal static Shape SafeThis(this Shape shape) {
            try {
                object testAccess = shape.Parent;
            }
            catch {
                return null;
            }
            return shape;
        }
    }
}
