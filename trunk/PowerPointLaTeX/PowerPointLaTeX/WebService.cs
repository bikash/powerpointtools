﻿#region Copyright Notice
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

namespace PowerPointLaTeX
{
    public class WebService : ILaTeXService
    {
        private struct URLData
        {
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
        private static URLData getURLData(string url)
        {
            HttpWebRequest request = (HttpWebRequest) HttpWebRequest.Create(url);
            request.Timeout = 3000;
            HttpWebResponse response = null;
            try {
                response = (HttpWebResponse) request.GetResponse();
            } catch {
                return new URLData();
            }

            Stream responseStream = response.GetResponseStream();

            Byte[] bytes = new Byte[response.ContentLength];
            int offset = 0;
            // apparently we need to everything packet by packet
            while (offset < response.ContentLength) {
                int numBytesRead = responseStream.Read(bytes, offset, (int) response.ContentLength - offset);
                offset += numBytesRead;
                if( numBytesRead == 0 ) {
                    response.Close();
                    return new URLData();
                }
            }

            response.Close();

            URLData data = new URLData();
            data.content = bytes.Take((int) response.ContentLength).ToArray();
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
            get { return "WebService using http://www.codecogs.com"; }
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
            Trace.Assert(System.Text.RegularExpressions.Regex.IsMatch(URLData.contentType, "gif|bmp|jpeg|png"));
            return URLData.content;
        }

        public System.Windows.Forms.UserControl GetPreferencesPage()
        {
            return null;
        }

        #endregion
    }
}
