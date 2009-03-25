using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PowerPointLaTeX
{
    internal class LaTeXServics
    {
        public ILaTeXService Service
        {
            get;
            private set;
        }

        public LaTeXServics()
        {
            Service = new WebService();
        }
    }
}
