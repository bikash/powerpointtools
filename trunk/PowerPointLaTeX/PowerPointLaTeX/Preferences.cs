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
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using PowerPointLaTeX.Properties;

namespace PowerPointLaTeX
{
    public partial class Preferences : Form
    {
        public Preferences() {
            InitializeComponent();

            // fill the serviceSelector
            serviceSelector.Items.Clear();
            serviceSelector.Items.AddRange( Globals.ThisAddIn.LaTeXServices.ServiceNames );

            Save();
        }

        private void Save() {
            Settings.Default.Save();
            MiKTexSettings.Default.Save();
        }

        private void Reload() {
            Settings.Default.Reload();
            MiKTexSettings.Default.Reload();

        }

        private void AbortButton_Click( object sender, EventArgs e ) {
            Reload();
        }

        private void OkButton_Click( object sender, EventArgs e ) {
            Save();
        }
    }
}
