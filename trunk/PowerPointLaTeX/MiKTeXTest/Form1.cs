using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace MiKTeXTest
{
    public partial class Form1 : Form
    {
        const string mikTexPath = @"d:\LaTeX\MiKTeX Portable";
        const string latexRelPath = @"miktex\bin\latex.exe";
        const string dvipngRelPath = @"miktex\bin\dvipng.exe";

        const string latexOptions = "-enable-installer -interaction=batchmode";
        const string dvipngOptions = "-T tight --depth --height -D 300 --noghostscript --picky -q -z 0";
        
        public Form1()
        {
            InitializeComponent();
        }

        private static string runMikTexBinary( string workingDir, string relPath, string arguments ) {
            Process process = new Process();
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.WorkingDirectory = workingDir;
            process.StartInfo.RedirectStandardOutput = true;
            process.StartInfo.CreateNoWindow = true;

            process.StartInfo.FileName = Path.Combine(mikTexPath, relPath);
            process.StartInfo.Arguments = arguments;
            process.Start();

            string output = process.StandardOutput.ReadToEnd();
            process.WaitForExit();

            return output;
        }

        private void runMiKTeX_Click(object sender, EventArgs e)
        {
            string tempTexFileName = Path.GetTempFileName();

            // write tex file
            FileStream texFileStream = new FileStream(tempTexFileName, FileMode.Truncate, FileAccess.Write);
            StreamWriter texFileOut = new StreamWriter(texFileStream);

            texFileOut.WriteLine(codeBox.Text);

            texFileOut.Close();

            String tempDir = Path.GetDirectoryName(tempTexFileName);
            String outputImagePath = Path.ChangeExtension( tempTexFileName, "png" );

            string latexOutput = runMikTexBinary( tempDir, latexRelPath, latexOptions + " \"" + tempTexFileName + "\"" );
            string dvipngOutput = runMikTexBinary( tempDir, dvipngRelPath, 
                dvipngOptions
                + " -o \"" + outputImagePath + "\""
                + " \"" + Path.ChangeExtension( tempTexFileName, "dvi" ) );

            latexOutputBox.Text = latexOutput;
            dvipngOutputBox.Text = dvipngOutput;

            int depth = Int32.Parse( Regex.Match( dvipngOutput, @"depth=(\d*)").Groups[1].Value );
            int height = Int32.Parse(Regex.Match(dvipngOutput, @"height=(\d*)").Groups[1].Value);

            depthInfo.Text = depth.ToString();
            heightInfo.Text = height.ToString();

            outputImage.Load(outputImagePath);

            // delete temp files
            string[] tempFiles = Directory.GetFiles(tempDir, Path.GetFileNameWithoutExtension(tempTexFileName) + ".*");
            foreach (string filePath in tempFiles)
            {
                File.Delete( filePath );
            }
        }
    }
}
