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
        private const string dvipngOptions = "-T tight --depth --height -D DPI --noghostscript --picky -q -z 0";

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

        private bool compileLatexCode(string code, int DPI, out byte[] imageData, out int baselineOffset) {
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

                // run it twice...
                // HACK: it wont work with changing dpi for some reason :( [5/23/2010 Andreas]
                string dvipngOutput = "";
                for( int i = 0; i < 2; i++ ) {
                    dvipngOutput = runConsoleProcess( tempDir, settings.DVIPNGPath,
                        dvipngOptions.Replace( "DPI", DPI.ToString() )
                        + " -o \"" + outputImagePath + "\""
                        + " \"" + Path.ChangeExtension( tempTexFileName, "dvi" ) );

                    lastLog += "\nDVIPNG Output:\n\n" + dvipngOutput;

                    if( File.Exists(outputImagePath) ) {
                        break;
                        // otherwise give it a second try - maybe METAFONT failed for some reason on the first attempt..
                    }
                }

                int depth = Int32.Parse( Regex.Match( dvipngOutput, @"depth=(\S*)" ).Groups[ 1 ].Value );
                int height = Int32.Parse( Regex.Match( dvipngOutput, @"height=(\S*)" ).Groups[ 1 ].Value );
                baselineOffset = depth;

                imageData = File.ReadAllBytes( outputImagePath );
            }
            // catch and ignore - something went wrong - oh well..
            catch {}
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

        public bool RenderLaTeXCode( string latexCode, out byte[] imageData, ref float pixelsPerEmHeight, out int baselineOffset )
        {
            // ignore pixelsPerEmHeight and simple use our fixed DPI value for now
            float latexFontSizePt = 10;
            float latexPrintPtPerInch = 72;
            int DPI = (int)( 0.5f + pixelsPerEmHeight / (latexFontSizePt / latexPrintPtPerInch) );
            if( DPI < 150 ) {
                DPI = 150;
            }
            // only allow steps of 10..
            DPI = DPI - DPI % 10;

            pixelsPerEmHeight = latexFontSizePt / latexPrintPtPerInch * DPI;

            string fullLatexCode = settings.MikTexTemplate.Replace( "LATEXCODE", Globals.ThisAddIn.Tool.ActivePresentation.SettingsTags().MiKTeXPreamble + "\n$" + latexCode + "$");

            lastLog = "";

            if( !compileLatexCode( fullLatexCode, DPI, out imageData, out baselineOffset ) ) {
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
