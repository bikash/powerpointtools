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
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Windows.Forms.Design;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLaTeX
{
    public partial class DeveloperTaskPaneControl : UserControl
    {
        public static string Title
        {
            get { return "LaTeX Developer Pane"; }
        }

        private Presentation Presentation
        {
            get
            {
                return Globals.ThisAddIn.Application.ActivePresentation;
            }
        }

        public DeveloperTaskPaneControl()
        {
            InitializeComponent();

            Globals.ThisAddIn.Application.WindowSelectionChange += new EApplication_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);
        }

        void Application_WindowSelectionChange(Selection Sel)
        {
            ProcessSelection(Sel);
        }

        private void ProcessSelection(Selection Sel)
        {
            tagsLayout.Controls.Clear();
            // an exception is thrown otherwise >_>
            if (Presentation.Final)
            {
                return;
            }
            switch (Sel.Type)
            {
                case PpSelectionType.ppSelectionShapes:
                case PpSelectionType.ppSelectionText:
                    AddShapesToSelection(Sel.ShapeRange);
                    break;
                case PpSelectionType.ppSelectionSlides:
                    AddSlidesToSelection(Sel.SlideRange);
                    break;
                case PpSelectionType.ppSelectionNone:
                    AddPresentationToSelection();
                    break;
            }
        }

        private void AddShapesToSelection(System.Collections.IEnumerable shapes)
        {
            foreach (Shape shape in shapes)
            {
                tagsLayout.Controls.Add(new TagsGrid(String.Format("Shape {0}:", shape.Id), shape.Tags));
            }
        }

        private void AddSlidesToSelection(System.Collections.IEnumerable slides)
        {
            foreach (Slide slide in slides)
            {
                tagsLayout.Controls.Add(new TagsGrid(String.Format("Shape {0} ({1}):", GetSlideTitle(slide), slide.SlideID), slide.Tags));
            }
        }

        private string GetSlideTitle(Slide slide)
        {
            string title = "<not available>";
            try
            {
                title = slide.Shapes.Title.TextFrame.TextRange.Text;
            }
            catch
            {
            }
            return title;
        }

        private void AddPresentationToSelection()
        {
            tagsLayout.Controls.Add(new TagsGrid(String.Format("Presentation {0}", Presentation.Name), Presentation.Tags));
        }

        private void refreshButton_Click(object sender, EventArgs e)
        {
            foreach(TagsGrid tagsGrid in tagsLayout.Controls) {
                tagsGrid.RefreshTags();
            }
        }

        private void selectAllButton_Click(object sender, EventArgs e)
        {
            tagsLayout.Controls.Clear();

            AddPresentationToSelection();
            AddSlidesToSelection(Presentation.Slides);
            foreach (Slide slide in Presentation.Slides)
            {
                AddShapesToSelection(slide.Shapes);
            }
        }

        private void useCurrentSelectionButton_Click(object sender, EventArgs e)
        {
            ProcessSelection(Globals.ThisAddIn.Application.ActiveWindow.Selection);
        }

    }
}
