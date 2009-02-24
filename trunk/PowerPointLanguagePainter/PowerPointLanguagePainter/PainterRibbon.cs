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
