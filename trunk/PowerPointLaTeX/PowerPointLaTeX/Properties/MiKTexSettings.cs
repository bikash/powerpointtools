using System;
using System.IO;
using System.Windows.Forms;

namespace PowerPointLaTeX.Properties {
    
    
    // This class allows you to handle specific events on the settings class:
    //  The SettingChanging event is raised before a setting's value is changed.
    //  The PropertyChanged event is raised after a setting's value is changed.
    //  The SettingsLoaded event is raised after the setting values are loaded.
    //  The SettingsSaving event is raised before the setting values are saved.
    internal sealed partial class MiKTexSettings {
        
        public MiKTexSettings() {
            this.PropertyChanged += this.PropertyChangedEventHandler;

            // To add event handlers for saving and changing settings, uncomment the lines below:
            //
            // this.SettingsSaving += this.SettingsSavingEventHandler;
            //
        }

        void PropertyChangedEventHandler(object sender, System.ComponentModel.PropertyChangedEventArgs e) {
            if (e.PropertyName.Equals("DistributionType") || e.PropertyName.Equals("DistributionPath")) {
                RecombinePaths();
            }
        }
        
        private void RecombinePaths() {
            // get the base path
            string relBasePath = "";
            switch ((MiKTeXService.Distribution)Enum.Parse(typeof(MiKTeXService.Distribution), DistributionType)) {
                case MiKTeXService.Distribution.MiKTeX:
                    relBasePath = MikTexBinPath;
                    break;
                case MiKTeXService.Distribution.TeXLive:
                    relBasePath = TexLiveBinPath;
                    break;
            }

            try {
                LatexPath = Path.Combine(DistributionPath, relBasePath, LatexFilename);
                DVIPNGPath = Path.Combine(DistributionPath, relBasePath, DVIPNGFilename);
            }
            catch (System.Exception ex) {
                MessageBox.Show(ex.ToString());
            }
        }
    }
}
