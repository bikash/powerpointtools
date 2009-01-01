using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;
using System.Diagnostics;
using Microsoft.Office.Tools;
using System.ComponentModel;

namespace PowerPointLaTeX
{
    public partial class ThisAddIn
    {
        internal LaTeXTool Tool
        {
            get;
            private set;
        }

        internal Settings Settings
        {
            get;
            private set;
        }

        internal CustomTaskPane DeveloperTaskPane;


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Tool = new LaTeXTool();
            Settings = new Settings();

            Settings.AutomaticCompilation = true;
            Settings.OfflineMode = false;

            DeveloperTaskPane = CustomTaskPanes.Add(new DeveloperTaskPaneControl(), DeveloperTaskPaneControl.Title);
            DeveloperTaskPane.Visible = Properties.Settings.Default.ShowDeveloperTaskPane;

            // register events
            Application.PresentationSave += new EApplication_PresentationSaveEventHandler(Application_PresentationSave);
            Application.SlideShowBegin += new EApplication_SlideShowBeginEventHandler(Application_SlideShowBegin);
            Application.WindowBeforeDoubleClick += new EApplication_WindowBeforeDoubleClickEventHandler(Application_WindowBeforeDoubleClick);
            Application.WindowSelectionChange += new EApplication_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);

            Properties.Settings.Default.PropertyChanged += new PropertyChangedEventHandler(Default_PropertyChanged);
            DeveloperTaskPane.VisibleChanged += new EventHandler(DeveloperTaskPane_VisibleChanged);
        }

        void DeveloperTaskPane_VisibleChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.ShowDeveloperTaskPane = DeveloperTaskPane.Visible;
        }

        void Default_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "ShowDeveloperTaskPane")
            {
                DeveloperTaskPane.Visible = (bool) ((Properties.Settings) sender)[e.PropertyName];
            }
        }

        private Dictionary<Presentation, Shape> oldTextShapeDict = new Dictionary<Presentation, Shape>();
        private Shape oldTextShape
        {
            get
            {
                if (!oldTextShapeDict.ContainsKey(Tool.ActivePresentation))
                {
                    return null;
                }
                return oldTextShapeDict[Tool.ActivePresentation];
            }
            set
            {
                oldTextShapeDict[Tool.ActivePresentation] = value;
            }
        }

        private void Application_WindowSelectionChange(Selection Sel)
        {

            // automatically select the parent (and thus all children) of a inline objects
            List<Shape> shapes = Sel.GetShapesFromShapeSelection();

            IEnumerable<Shape> parentShapes =
                from shape in shapes
                where shape.LaTeXTags().Type == LaTeXTool.EquationType.Inline
                select Tool.GetParentShape(shape);
            IEnumerable<Shape> shapeSuperset =
                from parentShape in parentShapes.Union(shapes)
                from inlineShape in Tool.GetInlineShapes(parentShape)
                select inlineShape;

            Sel.SelectShapes(parentShapes, false);
            Sel.SelectShapes(shapeSuperset, false);

            Shape textShape = Sel.GetShapeFromTextSelection();
            // recompile the old shape if necessary (do nothing if we click around in the same text shape though)
            if (!Settings.OfflineMode)
            {
                if (Settings.AutomaticCompilation)
                {
                    if (oldTextShape != null && oldTextShape != textShape)
                    {
                        Tool.CompileShape(oldTextShape.GetSlide(), oldTextShape);
                    }
                    if (textShape != null)
                    {
                        Tool.DecompileShape(Tool.ActiveSlide, textShape);
                    }
                }
            } else {
                if( textShape.LaTeXTags().Type == LaTeXTool.EquationType.HasCompiledInlines ) {
                    textShape = null;
                    Sel.Unselect();
                }
            }
            oldTextShape = textShape;
        }

        void Application_WindowBeforeDoubleClick(Selection Sel, ref bool Cancel)
        {
            // kill the MS engineers - kill them all and torture them slowly to death..
            // http://www.eggheadcafe.com/software/aspnet/33533167/ppt--windowbeforedoublec.aspx
        }

        void Application_SlideShowBegin(SlideShowWindow Wn)
        {
            // check whether anything still needs to be compiled and ask if necessary
            //throw new NotImplementedException();
        }

        void Application_PresentationSave(Presentation presentation)
        {
            // purge unused items from the cache to keep it smaller (thats the idea)
            presentation.CacheTags().PurgeUnused();
            //throw new NotImplementedException();
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
