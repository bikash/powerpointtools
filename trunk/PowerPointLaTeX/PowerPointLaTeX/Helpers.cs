using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PowerPointLaTeX
{
    class Helpers
    {
        static public bool IsEscapeCode(string code)
        {
            return code == "!";
        }
    }
}
