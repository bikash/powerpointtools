using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLaTeX
{
    class CacheTags
    {
        private Presentation presentation;

        public CacheTags(Presentation presentation)
        {
            this.presentation = presentation;
        }

        public void PurgeAll()
        {
            presentation.PurgeTags("IsCached#");
            presentation.PurgeTags("Cache#");
        }

        private bool IsCached(string code)
        {
            return "true" == presentation.GetTag("IsCached#" + code);
        }

        private byte[] GetCacheEntryFor(string code)
        {
            if (!IsCached(code))
            {
                return null;
            }

            byte[] data = Convert.FromBase64String(presentation.GetTag("Cache#" + code));
            return data;
        }

        private void SetCacheEntryFor(string code, byte[] data)
        {
            if (data != null)
            {
                presentation.SetTag("IsCached#" + code, "true");
                string base64 = Convert.ToBase64String(data);
                presentation.SetTag("Cache#" + code, base64);
            }
            else
            {
                presentation.ClearTag("IsCached#" + code);
                presentation.ClearTag("Cache#" + code);
            }
        }

        public byte[] this[string code]
        {
            get { return GetCacheEntryFor(code); }
            set { SetCacheEntryFor(code, value); }
        }
    }

}
