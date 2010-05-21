using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPointLaTeX.Properties;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using System.Drawing;

namespace PowerPointLaTeX
{
    public class MiKTeXService : ILaTeXService
    {
        private const string latexOptions = "-enable-installer -interaction=nonstopmode";
        private const string dvipngOptions = "-T tight --depth --height -D 300 --noghostscript --picky -q -z 0";

        private string lastLog;

        private MiKTexSettings settings {
            get { return MiKTexSettings.Default; }
        }

        private static string runConsoleProcess( string workingDir, string path, string arguments ) {
            Process process = new Process();
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.WorkingDirectory = workingDir;
            process.StartInfo.RedirectStandardOutput = true;
            process.StartInfo.CreateNoWindow = true;

            process.StartInfo.FileName = path;
            process.StartInfo.Arguments = arguments;
            process.Start();

            string output = process.StandardOutput.ReadToEnd();
            process.WaitForExit();

            return output;
        }

        private bool compileLatexCode(string code, out byte[] imageData, out int baselineOffset) {
            baselineOffset = 0;
            imageData = null;

            string tempTexFileName = Path.GetTempFileName();
            String tempDir = Path.GetDirectoryName( tempTexFileName );

            try {
                // write tex file
                FileStream texFileStream = new FileStream( tempTexFileName, FileMode.Truncate, FileAccess.Write );
                StreamWriter texFileOut = new StreamWriter( texFileStream );

                texFileOut.WriteLine( code );

                texFileOut.Close();

                String outputImagePath = Path.ChangeExtension( tempTexFileName, "png" );

                string latexOutput = runConsoleProcess( tempDir, settings.LatexPath, latexOptions + " \"" + tempTexFileName + "\"" );

                lastLog += "Latex Output:\n\n" + latexOutput;

                string dvipngOutput = runConsoleProcess( tempDir, settings.DVIPNGPath,
                    dvipngOptions
                    + " -o \"" + outputImagePath + "\""
                    + " \"" + Path.ChangeExtension( tempTexFileName, "dvi" ) );

                lastLog += "\nDVIPNG Output:\n\n" + dvipngOutput;

                int depth = Int32.Parse( Regex.Match( dvipngOutput, @"depth=(\S*)" ).Groups[ 1 ].Value );
                int height = Int32.Parse( Regex.Match( dvipngOutput, @"height=(\S*)" ).Groups[ 1 ].Value );
                baselineOffset = depth;

                imageData = File.ReadAllBytes( outputImagePath );
            }
            finally {
                // delete temp files
                string[] tempFiles = Directory.GetFiles( tempDir, Path.GetFileNameWithoutExtension( tempTexFileName ) + ".*" );
                foreach( string filePath in tempFiles ) {
                    File.Delete( filePath );
                }
            }

            return (imageData != null);
        }

        #region ILaTeXService Members

        public string AboutNotice
        {
            get { return "Andreas Kirsch 2010"; }
        }

        public string SeriveName
        {
            get { return "MiKTeX Service"; }
        }

        public bool RenderLaTeXCode( string latexCode, out byte[] imageData, out int baselineOffset )
        {
            string fullLatexCode = settings.MikTexTemplate.Replace( "LATEXCODE", "$" + latexCode + "$");

            lastLog = "";

            if( !compileLatexCode( fullLatexCode, out imageData, out baselineOffset ) ) {
                imageData = null;
                return false;
            }

            return true;
        }

        public string GetLastErrorReport()
        {
            return lastLog;
        }

        #endregion
    }
}
