using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PowerPointLaTeX {
    class NullService : ILaTeXRenderingService {
        #region ILaTeXService Members

        public string AboutNotice {
            get { return "Choose this to only use the presentation cache."; }
        }

        public string SeriveName {
            get { return "Cache Only"; }
        }

        public bool RenderLaTeXCode( string latexCode, out byte[] imageData, ref float pixelsPerEmHeight, out int baselineOffset ) {
            imageData = null;
            pixelsPerEmHeight = 0;
            baselineOffset = 0;
            return false;
        }

        public string GetLastErrorReport() {
            return "No LaTeX service chosen!";
        }

        #endregion
    }
}
