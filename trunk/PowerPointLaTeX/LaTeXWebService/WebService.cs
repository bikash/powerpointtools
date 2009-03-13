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

namespace LaTeXWebService
{
    public class WebService
    {
        public struct URLData {
            public byte[] content;
            public string contentType;
        }

        private const string webServiceURL = @"http://www.codecogs.com/png.latex?\bg_white&space;\300dpi&space;"; // "http://l.wordpress.com/latex.php?bg=ffffff&fg=000000&latex=";

        private static string getRequestURL(string latexCode)
        {
            return webServiceURL + HttpUtility.HtmlEncode(latexCode);
        }

        /// <summary>
        /// Read the url
        /// </summary>
        /// <param name="url"></param>
        /// <returns>Null if not everything could be read</returns>
        private static URLData getURLData(string url) {
            WebRequest request = HttpWebRequest.Create(url);
            request.Timeout = 3000;
            WebResponse response = request.GetResponse();

            Stream responseStream = response.GetResponseStream();

            Byte[] bytes = new Byte[response.ContentLength];
            int numBytesRead = responseStream.Read(bytes, 0, (int)response.ContentLength);

            // just return null if we can't read the whole packet for some reason (e.g. connection drop or similar)
            if( numBytesRead != response.ContentLength ) {
                return new URLData();
            }

            URLData data = new URLData();
            data.content = bytes.Take(numBytesRead).ToArray();
            data.contentType = response.ContentType;
            return data;
        }

        public static URLData compileLaTeX(string latexCode)
        {
            return getURLData( getRequestURL(latexCode) );
        }
    }
}
