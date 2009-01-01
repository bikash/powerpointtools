using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PowerPointLaTeX
{
    class Settings
    {
        public delegate void ToggleChangedEventHandler(bool isChecked);
        public event ToggleChangedEventHandler onAutomaticCompilationChanged = null;
        public event ToggleChangedEventHandler onOfflineModeChanged = null;

        internal bool AutomaticCompilation
        {
            get { return Globals.Ribbons.LaTeXRibbon.AutomaticCompilationToggle.Checked; }
            set { 
                Globals.Ribbons.LaTeXRibbon.AutomaticCompilationToggle.Checked = value;
                ToggleChangedEventHandler handler = onAutomaticCompilationChanged;
                if( handler != null ) {
                    handler(value);
                }
            }
        }

        internal bool OfflineMode
        {
            get { return Globals.Ribbons.LaTeXRibbon.PresentationModeToggle.Checked; }
            set {
                Globals.Ribbons.LaTeXRibbon.PresentationModeToggle.Checked = value;
                ToggleChangedEventHandler handler = onOfflineModeChanged;
                if( handler != null ) {
                    handler(value);
                }
            }
        }

    }
}
