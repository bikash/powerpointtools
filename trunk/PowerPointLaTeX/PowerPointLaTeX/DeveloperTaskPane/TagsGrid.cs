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
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLaTeX
{
    public partial class TagsGrid : UserControl
    {
        Tags tags;

        public TagsGrid(string name, Tags tags)
        {
            InitializeComponent();

            itemName.Text = name;
            this.tags = tags;

            this.Dock = DockStyle.Fill;
            this.AutoSize = true;

            RefreshTags();
        }

        public void RefreshTags()
        {
            tagsGridView.Rows.Clear();

            for (int i = 1; i <= tags.Count; i++)
            {
                tagsGridView.Rows.Add(new string[] { tags.Name(i), tags.Value(i) });
            }
        }
    }
}
