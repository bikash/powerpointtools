#region Copyright Notice
// This file is part of PowerPoint LaTeX.
// 
// Copyright (C) 2008/2009 Andreas Kirsch
// 
// PowerPoint LaTeX is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
// 
// PowerPoint LaTeX is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
// 
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <http://www.gnu.org/licenses/>.
#endregion

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.Web;
using System.IO;
using System.Diagnostics;
using System.Drawing;

namespace PowerPointLaTeX
{
    public class WebService : ILaTeXRenderingService
    {
        private struct URLData
        {
            public byte[] content;
            public string contentType;
        }

        private const string webServiceURL = @"http://www.codecogs.com/png.latex?\bg_white \300dpi "; // "http://l.wordpress.com/latex.php?bg=ffffff&fg=000000&latex=";

        private static string getRequestURL(string latexCode)
        {
            return webServiceURL + latexCode;
        }

        /// <summary>
        /// Read the url
        /// </summary>
        /// <param name="url"></param>
        /// <returns>Null if not everything could be read</returns>
        private static URLData getURLData(string url)
        {
            // TODO: clean this code up *pretty please* (but make sure this instable PoS doesnt break) [5/13/2009 Andreas]
            HttpWebRequest request = (HttpWebRequest) WebRequest.Create(url);
            request.Timeout = 6000;
            request.Method = "GET";
            HttpWebResponse response = null;
            try {
                response = (HttpWebResponse) request.GetResponse();
            } catch {
                return new URLData();
            }

            Stream responseStream = response.GetResponseStream();
            byte[] buffer = new byte[1024*512];
            int offset = 0;
            // apparently we need to everything packet by packet
            while (true) {
                int numBytesRead = responseStream.Read(buffer, offset, 2048);
                offset += numBytesRead;
                if( numBytesRead == 0) {
                    if (offset == 0) {
                        response.Close();
                        return new URLData();
                    }
                    else {
                        break;
                    }
                }
            }

            response.Close();

            URLData data = new URLData();
            data.content = buffer.Take(offset).ToArray();
            data.contentType = response.ContentType;
            return data;
        }

        /// <summary>
        /// Always returns a valid URLData object
        /// </summary>
        /// <param name="latexCode"></param>
        /// <returns></returns>
        private static URLData compileLaTeX(string latexCode)
        {
            return getURLData(getRequestURL(latexCode));
        }

        #region ILaTeXService Members

        public string AboutNotice
        {
            get { return "Uses the LaTeX WebService from http://www.codecogs.com."; }
        }

        public string SeriveName
        {
            get { return "LaTeX WebService"; }
        }

        public byte[] GetImageDataForLaTeXCode(string latexCode)
        {
            URLData URLData = compileLaTeX(latexCode);
            // TODO: replace all the null checks with exception handling? [2/26/2009 Andreas]
            if (URLData.content == null)
            {
                return null;
            }
            Trace.Assert(System.Text.RegularExpressions.Regex.IsMatch(URLData.contentType, "gif|bmp|jpeg|png"), "Unknown image format returned from the web service");
            return URLData.content;
        }

        public bool RenderLaTeXCode(string latexCode, out byte[] imageData, ref float pixelsPerEmHeight, out int baselineOffset) {
            // measured value
            pixelsPerEmHeight = 59;
 
            // baseline offset not supported!
            baselineOffset = 0;

            URLData URLData = compileLaTeX(latexCode);
            // TODO: replace all the null checks with exception handling? [2/26/2009 Andreas]
            if (URLData.content == null)
            {
                imageData = null;
                return false;
            }
            Trace.Assert(System.Text.RegularExpressions.Regex.IsMatch(URLData.contentType, "gif|bmp|jpeg|png"));

            imageData = URLData.content;
            return true;
        }

        public string GetLastErrorReport() {
            return "No information provided.";
        }

        public float DPI {
            get {
                return 300;
            }
        }

        #endregion
    }
}
