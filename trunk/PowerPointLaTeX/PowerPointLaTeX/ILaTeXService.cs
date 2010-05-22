using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Windows.Forms;

namespace PowerPointLaTeX
{
    public interface ILaTeXService
    {
        string AboutNotice {
            get;
        }

        string SeriveName {
            get;
        }

        /// <summary>
        /// Get the raw data of an image that can be read 
        /// </summary>
        /// <param name="latexCode"></param>
        /// <param name="image">the actual image of the rendered latexCode</param>
        /// <param name="baselineOffset"> (from the dvipng manpage)
        /// It reports the number of pixels from the bottom of the image to the baseline of the image.
        /// This can be used for vertical positioning of the image in, e.g., web documents, where one would use (Cascading StyleSheets 1)
        /// The depth is a negative offset in this case, so the minus sign is necessary, and the unit is pixels (px).
        /// </param>
        /// <returns>returns false if there was an error</returns>
        bool RenderLaTeXCode(string latexCode, out byte[] imageData, ref float pixelsPerEmHeight, out int baselineOffset);

        string GetLastErrorReport();
    }
}
