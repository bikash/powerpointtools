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
        /// <returns>null if there was an error or something similar</returns>
        byte[] GetImageDataForLaTeXCode(string latexCode);

        UserControl GetPreferencesPage();
    }
}
