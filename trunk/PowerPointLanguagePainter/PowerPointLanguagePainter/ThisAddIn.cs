#region Copyright Notice
// This file is part of PowerPoint Language Painter.
// 
// Copyright (C) 2008/2009 Andreas Kirsch
// 
// PowerPoint Language Painter is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
// 
// PowerPoint Language Painter is distributed in the hope that it will be useful,
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
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLanguagePainter
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.WindowSelectionChange += new EApplication_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);
        }

        void Application_WindowSelectionChange(Selection Sel)
        {
            Presentation presentation = Application.ActivePresentation;
            if (presentation.Final)
            {
                return;
            }
            if (Properties.Settings.Default.EnablePainting)
            {
                // this paints only words
                /*                
                TextRange textRange = Sel.TextRange;
                                if( textRange != null ) {
                                    textRange.LanguageID = Properties.Settings.Default.LanguageID;
                                }*/
                if (Sel.Type == PpSelectionType.ppSelectionText)
                {
                    Shape parent = Sel.ShapeRange[1];
                    if (parent.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue && parent.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                    {
                        parent.TextFrame.TextRange.LanguageID = Properties.Settings.Default.LanguageID;
                    }
                }
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Properties.Settings.Default.Save();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
