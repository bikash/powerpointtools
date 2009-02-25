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

        private static URLData getURLData(string url) {
            WebRequest request = HttpWebRequest.Create(url);
            request.Timeout = 3000;
            WebResponse response = request.GetResponse();

            Stream responseStream = response.GetResponseStream();

            Byte[] bytes = new Byte[response.ContentLength];
            int numBytesRead = responseStream.Read(bytes, 0, (int)response.ContentLength);

            //Trace.Assert(numBytesRead == response.ContentLength);

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
