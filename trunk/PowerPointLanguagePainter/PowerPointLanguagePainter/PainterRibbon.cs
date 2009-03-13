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
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Core;

namespace PowerPointLanguagePainter
{
    public partial class PainterRibbon : OfficeRibbon
    {
        public PainterRibbon()
        {
            InitializeComponent();

            Properties.Settings.Default.PropertyChanged += new System.ComponentModel.PropertyChangedEventHandler(Default_PropertyChanged);

            UpdateLanguage();
            paintToggleButton.Checked = Properties.Settings.Default.EnablePainting;
        }

        void Default_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "EnablePainting")
            {
                paintToggleButton.Checked = Properties.Settings.Default.EnablePainting;
            }
            else if (e.PropertyName == "LanguageID")
            {
                UpdateLanguage();
            }
        }

        private void UpdateLanguage()
        {
            paintToggleButton.Label = "Paint " + GetStringFromLanguageID(Properties.Settings.Default.LanguageID);
        }

        private string GetStringFromLanguageID(MsoLanguageID languageID)
        {
            switch (languageID)
            {
                case MsoLanguageID.msoLanguageIDGerman:
                    return "German";
                case MsoLanguageID.msoLanguageIDEnglishUS:
                    return "English (US)";
                case MsoLanguageID.msoLanguageIDFrench:
                    return "French";
                case MsoLanguageID.msoLanguageIDSpanish:
                    return "Spanish";
            }
            return languageID.ToString();
        }
        
        // TODO: lots of duplicate one-liners - look into merging them [2/24/2009 Andreas]
        private void languageIDEnglishUS_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.LanguageID = MsoLanguageID.msoLanguageIDEnglishUS;
        }

        private void languageIDGerman_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.LanguageID = MsoLanguageID.msoLanguageIDGerman;
        }

        private void languageIDFrench_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.LanguageID = MsoLanguageID.msoLanguageIDFrench;
        }

        private void languageIDSpanish_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.LanguageID = MsoLanguageID.msoLanguageIDSpanish;
        }

        private void paintToggleButton_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.EnablePainting = paintToggleButton.Checked;
        }
    }
}
