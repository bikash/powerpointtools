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
using System.IO;
using System.Drawing;
using System.Diagnostics;
using System.Windows.Forms;
using System.Threading;

namespace PowerPointLaTeX
{
    public enum EquationType
    {
        None,
        // has inline LaTeX codes (but not compiled)
        HasInlines,
        // has compiled LaTeX codes
        HasCompiledInlines,
        // a compiled LaTeX code element (picture)
        Inline,
        // an equation (picture)
        Equation,
    }

    static class EquationTypeShapeExtension
    {
        public static bool IsEquation(this Shape shape)
        {
            return shape.LaTeXTags().Type == EquationType.Equation;
        }
    }

    /// <summary>
    /// Contains all the important methods, etc.
    /// Instantiated by the add-in
    /// </summary>
    static class LaTeXTool
    {
        static private Microsoft.Office.Interop.PowerPoint.Application Application
        {
            get
            {
                return Globals.ThisAddIn.Application;
            }
        }

        static internal Presentation ActivePresentation
        {
            get { return Application.ActivePresentation; }
        }

        static internal Slide ActiveSlide
        {
            get { return Application.ActiveWindow.View.Slide as Slide; }
        }

        /// <summary>
        /// returns whether the addin is enabled in the current context (ie presentation)
        /// but is also affected by the global addin settings, of course.
        /// </summary>
        static internal bool AddInEnabled
        {
            get
            {
                return !ActivePresentation.Final && Properties.Settings.Default.EnableAddIn && Compatibility.IsSupportedPresentation(ActivePresentation);
            }
        }

        static private void FinalizeShape(Slide slide, Shape shape)
        {
            InlineFormulas.Embedding.CompileShape(slide, shape);
            shape.LaTeXTags().Clear();
        }

        static public void CompileSlide(Slide slide)
        {
            ShapeWalker.WalkSlide(slide, InlineFormulas.Embedding.CompileShape);
        }

        static public void DecompileSlide(Slide slide)
        {
            ShapeWalker.WalkSlide(slide, InlineFormulas.Embedding.DecompileShape);
        }

        static public void CompilePresentation(Presentation presentation)
        {
            ShapeWalker.WalkPresentation(presentation, InlineFormulas.Embedding.CompileShape);
        }

        /// <summary>
        /// Removes all tags and all pictures that belong to inline formulas
        /// </summary>
        /// <param name="slide"></param>
        static public void FinalizePresentation(Presentation presentation)
        {
            ShapeWalker.WalkPresentation(presentation, FinalizeShape);
            // purge the cache, too
            presentation.CacheTags().PurgeAll();
            presentation.SettingsTags().Clear();
        }
    }
}
