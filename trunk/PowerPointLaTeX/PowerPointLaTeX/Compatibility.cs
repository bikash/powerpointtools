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
using System.Windows.Forms;
using System.Diagnostics;

namespace PowerPointLaTeX
{
    static class Compatibility
    {
        static public int StorageVersion = 1;

        static private Dictionary<Presentation, bool> PresentationSupported = new Dictionary<Presentation, bool>();

        static public void Application_PresentationOpen(Presentation presentation)
        {
            bool supportedFormat = internalIsSupportedPresentation(presentation);
            PresentationSupported.Add(presentation, supportedFormat );

            if (!supportedFormat)
            {
                // FIXME: this is duplicated code from Preferences.cs [9/13/2010 Andreas]
                var assName = System.Reflection.Assembly.GetExecutingAssembly().GetName();
                string appInfo = assName.Name + " " + assName.Version;
                MessageBox.Show("Presentation '" + presentation.Name + "' uses an unsupported storage format version: " + GetStorageVersion(presentation) + "). " + appInfo + " only supports version " + StorageVersion + ". Please upgrade.", assName.Name);
            }
            else
            {
                UpgradePresentation(presentation);
            }
        }

        static public void Application_PresentationNew(Presentation presentation)
        {
            PresentationSupported.Add(presentation, true);
            SetStorageVersion(presentation, StorageVersion);
        }

        static public void Application_PresentationClose(Presentation presentation) {
            // PowerPoint 2007 raises this event before showing the "Save" dialog to the user, where the user can still cancel closing the current presentation
            // this obviously breaks the intention of the code below
            // "Solution": Don't remove the presentation - ie "leak memory" in case the user decides against closing the document
            // TODO: another solution would be to implement a cache policy and add the presentation back if it is not found in IsSupportedPresentation [9/13/2010 Andreas]
            //PresentationSupported.Remove(presentation);
        }

        static private int GetStorageVersion( Presentation presentation ) {
            AddInTagInt storageVersion = new AddInTagInt( presentation.Tags, "StorageVersion");

            return storageVersion;
        }

        static private void SetStorageVersion(Presentation presentation, int versionNumber )
        {
            AddInTagInt storageVersion = new AddInTagInt(presentation.Tags, "StorageVersion");

            storageVersion.value = versionNumber;
        }

        static private bool internalIsSupportedPresentation(Presentation presentation)
        {
            return GetStorageVersion(presentation) < StorageVersion;
        }

        static public bool IsSupportedPresentation( Presentation presentation ) {
            bool supported;
            if( PresentationSupported.TryGetValue(presentation, out supported) )
                return supported;

            // this shouldnt happen because every presentation should be added to PresentationSupported..
            Debug.Assert(false);
            return internalIsSupportedPresentation(presentation);
        }

        static private void UpgradePresentation( Presentation presentation ) {
            int oldVersion = GetStorageVersion(presentation);
            if( oldVersion < 1 ) {
                UpgradePresentation_0_1(presentation);
            }

            SetStorageVersion(presentation, StorageVersion);
        }

        static private void UpgradePresentation_0_1( Presentation presentation ) {
            // update cache tags
            // (nothing to do)

            // update settings tags
            const string MikTexTemplateContent =
@"\documentclass{article} 
\usepackage{amsmath}
\usepackage{amsthm}
\usepackage{amssymb}
\usepackage[active,displaymath,textmath,tightpage]{preview}
\usepackage{bm}

\begin{document}
\begin{preview}

LATEXCODE

\end{preview}
\end{document}
";
            AddInTagString MiKTeXPreamble = new AddInTagString(presentation.Tags, "MiKTeXPreamble");
            presentation.SettingsTags().MiKTeXTemplate.value =
                MikTexTemplateContent.Replace("LATEXCODE", MiKTeXPreamble + "\r\n$LATEXCODE$");
            MiKTeXPreamble.Clear();

            // update shape tags
            ShapeWalker.WalkPresentation(presentation,
                    delegate(Slide slide, Shape shape)
                    {
                        LaTeXTags latexTags = shape.LaTeXTags();
                        // skip if it's a shape that doesn't need to be updated
                        if( latexTags.Type != EquationType.Inline && latexTags.Type != EquationType.Equation ) {
                            return;
                        }

                        // set the default font size
                        if (latexTags.FontSize == 0)
                        {
                            latexTags.FontSize.value = 36;
                        }

                        // remove unused tag
                        (new AddInTagFloat(shape.Tags, "PixelsPerEmHeight")).Clear();
                    }
                );
        }
    }
}
