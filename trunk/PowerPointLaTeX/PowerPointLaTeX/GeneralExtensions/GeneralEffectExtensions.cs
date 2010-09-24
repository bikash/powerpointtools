using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLaTeX
{
    static class GeneralEffectExtensions
    {
        public static int GetSafeParagraph(this Effect effect)
        {
            try
            {
                return effect.Paragraph;
            }
            catch
            {
                return 1;
            }
        }
    }
}
