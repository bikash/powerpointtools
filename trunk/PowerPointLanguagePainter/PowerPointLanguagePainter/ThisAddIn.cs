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
