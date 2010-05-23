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

namespace PowerPointLaTeX
{
    partial class Preferences
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.Label label1;
            System.Windows.Forms.GroupBox groupBox1;
            System.Windows.Forms.Label label2;
            System.Windows.Forms.Label label3;
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager( typeof( Preferences ) );
            this.miktexPathBrowserButton = new System.Windows.Forms.Button();
            this.miktexPathBox = new System.Windows.Forms.TextBox();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.tabControl = new System.Windows.Forms.TabControl();
            this.generalPage = new System.Windows.Forms.TabPage();
            this.mikTexOptions = new System.Windows.Forms.TabPage();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.miktexPreambleBox = new System.Windows.Forms.TextBox();
            this.webServiceOptions = new System.Windows.Forms.TabPage();
            this.aboutPage = new System.Windows.Forms.TabPage();
            this.aboutBox = new System.Windows.Forms.RichTextBox();
            this.OkButton = new System.Windows.Forms.Button();
            this.AbortButton = new System.Windows.Forms.Button();
            this.miktexPathBrowser = new System.Windows.Forms.FolderBrowserDialog();
            this.serviceSelector = new System.Windows.Forms.ComboBox();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            label1 = new System.Windows.Forms.Label();
            groupBox1 = new System.Windows.Forms.GroupBox();
            label2 = new System.Windows.Forms.Label();
            label3 = new System.Windows.Forms.Label();
            groupBox1.SuspendLayout();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.tabControl.SuspendLayout();
            this.generalPage.SuspendLayout();
            this.mikTexOptions.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.aboutPage.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new System.Drawing.Point( 20, 15 );
            label1.Name = "label1";
            label1.Size = new System.Drawing.Size( 103, 13 );
            label1.TabIndex = 0;
            label1.Text = "Use LaTeX Service:";
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add( this.miktexPathBrowserButton );
            groupBox1.Controls.Add( this.miktexPathBox );
            groupBox1.Controls.Add( label2 );
            groupBox1.Location = new System.Drawing.Point( 8, 6 );
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new System.Drawing.Size( 416, 45 );
            groupBox1.TabIndex = 0;
            groupBox1.TabStop = false;
            groupBox1.Text = "Paths";
            // 
            // miktexPathBrowserButton
            // 
            this.miktexPathBrowserButton.Location = new System.Drawing.Point( 382, 11 );
            this.miktexPathBrowserButton.Name = "miktexPathBrowserButton";
            this.miktexPathBrowserButton.Size = new System.Drawing.Size( 28, 23 );
            this.miktexPathBrowserButton.TabIndex = 2;
            this.miktexPathBrowserButton.Text = "...";
            this.miktexPathBrowserButton.UseVisualStyleBackColor = true;
            this.miktexPathBrowserButton.Click += new System.EventHandler( this.miktexPathBrowserButton_Click );
            // 
            // miktexPathBox
            // 
            this.miktexPathBox.Location = new System.Drawing.Point( 85, 13 );
            this.miktexPathBox.Name = "miktexPathBox";
            this.miktexPathBox.Size = new System.Drawing.Size( 297, 20 );
            this.miktexPathBox.TabIndex = 1;
            this.miktexPathBox.TextChanged += new System.EventHandler( this.miktexPathBox_TextChanged );
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new System.Drawing.Point( 6, 16 );
            label2.Name = "label2";
            label2.Size = new System.Drawing.Size( 73, 13 );
            label2.TabIndex = 0;
            label2.Text = "MiKTeX Path:";
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new System.Drawing.Point( 7, 20 );
            label3.Name = "label3";
            label3.Size = new System.Drawing.Size( 343, 13 );
            label3.TabIndex = 0;
            label3.Text = "Preamble (inserted before all formulas - use this to add new commands):";
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel2;
            this.splitContainer1.Location = new System.Drawing.Point( 0, 0 );
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add( this.tabControl );
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add( this.OkButton );
            this.splitContainer1.Panel2.Controls.Add( this.AbortButton );
            this.splitContainer1.Size = new System.Drawing.Size( 440, 299 );
            this.splitContainer1.SplitterDistance = 270;
            this.splitContainer1.TabIndex = 0;
            // 
            // tabControl
            // 
            this.tabControl.Controls.Add( this.generalPage );
            this.tabControl.Controls.Add( this.mikTexOptions );
            this.tabControl.Controls.Add( this.webServiceOptions );
            this.tabControl.Controls.Add( this.aboutPage );
            this.tabControl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl.Location = new System.Drawing.Point( 0, 0 );
            this.tabControl.Multiline = true;
            this.tabControl.Name = "tabControl";
            this.tabControl.SelectedIndex = 0;
            this.tabControl.Size = new System.Drawing.Size( 440, 270 );
            this.tabControl.TabIndex = 1;
            // 
            // generalPage
            // 
            this.generalPage.Controls.Add( this.checkBox1 );
            this.generalPage.Controls.Add( this.serviceSelector );
            this.generalPage.Controls.Add( label1 );
            this.generalPage.Location = new System.Drawing.Point( 4, 22 );
            this.generalPage.Name = "generalPage";
            this.generalPage.Size = new System.Drawing.Size( 432, 244 );
            this.generalPage.TabIndex = 2;
            this.generalPage.Text = "General";
            this.generalPage.UseVisualStyleBackColor = true;
            // 
            // mikTexOptions
            // 
            this.mikTexOptions.Controls.Add( this.groupBox2 );
            this.mikTexOptions.Controls.Add( groupBox1 );
            this.mikTexOptions.Location = new System.Drawing.Point( 4, 22 );
            this.mikTexOptions.Name = "mikTexOptions";
            this.mikTexOptions.Padding = new System.Windows.Forms.Padding( 3 );
            this.mikTexOptions.Size = new System.Drawing.Size( 432, 244 );
            this.mikTexOptions.TabIndex = 0;
            this.mikTexOptions.Text = "MiKTeX Service";
            this.mikTexOptions.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add( this.miktexPreambleBox );
            this.groupBox2.Controls.Add( label3 );
            this.groupBox2.Location = new System.Drawing.Point( 8, 57 );
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size( 416, 181 );
            this.groupBox2.TabIndex = 1;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Presentation Settings";
            // 
            // miktexPreambleBox
            // 
            this.miktexPreambleBox.Location = new System.Drawing.Point( 7, 37 );
            this.miktexPreambleBox.Multiline = true;
            this.miktexPreambleBox.Name = "miktexPreambleBox";
            this.miktexPreambleBox.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.miktexPreambleBox.Size = new System.Drawing.Size( 403, 138 );
            this.miktexPreambleBox.TabIndex = 1;
            // 
            // webServiceOptions
            // 
            this.webServiceOptions.Location = new System.Drawing.Point( 4, 22 );
            this.webServiceOptions.Name = "webServiceOptions";
            this.webServiceOptions.Size = new System.Drawing.Size( 432, 244 );
            this.webServiceOptions.TabIndex = 3;
            this.webServiceOptions.Text = "Web Service";
            this.webServiceOptions.UseVisualStyleBackColor = true;
            // 
            // aboutPage
            // 
            this.aboutPage.Controls.Add( this.aboutBox );
            this.aboutPage.Location = new System.Drawing.Point( 4, 22 );
            this.aboutPage.Name = "aboutPage";
            this.aboutPage.Padding = new System.Windows.Forms.Padding( 3 );
            this.aboutPage.Size = new System.Drawing.Size( 432, 244 );
            this.aboutPage.TabIndex = 1;
            this.aboutPage.Text = "About..";
            this.aboutPage.UseVisualStyleBackColor = true;
            // 
            // aboutBox
            // 
            this.aboutBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.aboutBox.Cursor = System.Windows.Forms.Cursors.Default;
            this.aboutBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.aboutBox.Location = new System.Drawing.Point( 3, 3 );
            this.aboutBox.Name = "aboutBox";
            this.aboutBox.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.None;
            this.aboutBox.Size = new System.Drawing.Size( 426, 238 );
            this.aboutBox.TabIndex = 0;
            this.aboutBox.Text = "INSERT_APP_INFO\n(PowerPoint Addin)\n\nby Andreas \'BlackHC\' Kirsch\n\nINSERT_ABOUT_SER" +
                "VICES\n\n";
            // 
            // OkButton
            // 
            this.OkButton.Anchor = ((System.Windows.Forms.AnchorStyles) ((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.OkButton.AutoSize = true;
            this.OkButton.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.OkButton.Location = new System.Drawing.Point( 271, -1 );
            this.OkButton.Name = "OkButton";
            this.OkButton.Size = new System.Drawing.Size( 80, 25 );
            this.OkButton.TabIndex = 1;
            this.OkButton.Text = "OK";
            this.OkButton.UseVisualStyleBackColor = true;
            this.OkButton.Click += new System.EventHandler( this.OkButton_Click );
            // 
            // AbortButton
            // 
            this.AbortButton.Anchor = ((System.Windows.Forms.AnchorStyles) ((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.AbortButton.AutoSize = true;
            this.AbortButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.AbortButton.Location = new System.Drawing.Point( 357, -1 );
            this.AbortButton.Name = "AbortButton";
            this.AbortButton.Size = new System.Drawing.Size( 80, 25 );
            this.AbortButton.TabIndex = 0;
            this.AbortButton.Text = "Cancel";
            this.AbortButton.UseVisualStyleBackColor = true;
            this.AbortButton.Click += new System.EventHandler( this.AbortButton_Click );
            // 
            // miktexPathBrowser
            // 
            this.miktexPathBrowser.RootFolder = System.Environment.SpecialFolder.MyComputer;
            this.miktexPathBrowser.ShowNewFolderButton = false;
            // 
            // serviceSelector
            // 
            this.serviceSelector.DataBindings.Add( new System.Windows.Forms.Binding( "Text", global::PowerPointLaTeX.Properties.Settings.Default, "LatexService", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged ) );
            this.serviceSelector.FormattingEnabled = true;
            this.serviceSelector.Location = new System.Drawing.Point( 129, 12 );
            this.serviceSelector.Name = "serviceSelector";
            this.serviceSelector.Size = new System.Drawing.Size( 123, 21 );
            this.serviceSelector.TabIndex = 1;
            this.serviceSelector.Text = global::PowerPointLaTeX.Properties.Settings.Default.LatexService;
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Checked = global::PowerPointLaTeX.Properties.Settings.Default.EnableAddIn;
            this.checkBox1.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox1.DataBindings.Add( new System.Windows.Forms.Binding( "Checked", global::PowerPointLaTeX.Properties.Settings.Default, "EnableAddIn", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged ) );
            this.checkBox1.Location = new System.Drawing.Point( 129, 50 );
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size( 90, 17 );
            this.checkBox1.TabIndex = 3;
            this.checkBox1.Text = "Enable AddIn";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // Preferences
            // 
            this.AcceptButton = this.OkButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF( 6F, 13F );
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.CancelButton = this.AbortButton;
            this.ClientSize = new System.Drawing.Size( 440, 299 );
            this.Controls.Add( this.splitContainer1 );
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow;
            this.Icon = ((System.Drawing.Icon) (resources.GetObject( "$this.Icon" )));
            this.Name = "Preferences";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "PowerPointLaTeX Preferences";
            groupBox1.ResumeLayout( false );
            groupBox1.PerformLayout();
            this.splitContainer1.Panel1.ResumeLayout( false );
            this.splitContainer1.Panel2.ResumeLayout( false );
            this.splitContainer1.Panel2.PerformLayout();
            this.splitContainer1.ResumeLayout( false );
            this.tabControl.ResumeLayout( false );
            this.generalPage.ResumeLayout( false );
            this.generalPage.PerformLayout();
            this.mikTexOptions.ResumeLayout( false );
            this.groupBox2.ResumeLayout( false );
            this.groupBox2.PerformLayout();
            this.aboutPage.ResumeLayout( false );
            this.ResumeLayout( false );

        }

        #endregion

        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.TabControl tabControl;
        private System.Windows.Forms.TabPage generalPage;
        private System.Windows.Forms.TabPage mikTexOptions;
        private System.Windows.Forms.TabPage aboutPage;
        private System.Windows.Forms.RichTextBox aboutBox;
        private System.Windows.Forms.Button OkButton;
        private System.Windows.Forms.Button AbortButton;
        private System.Windows.Forms.ComboBox serviceSelector;
        private System.Windows.Forms.TabPage webServiceOptions;
        private System.Windows.Forms.TextBox miktexPathBox;
        private System.Windows.Forms.FolderBrowserDialog miktexPathBrowser;
        private System.Windows.Forms.Button miktexPathBrowserButton;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.TextBox miktexPreambleBox;
        private System.Windows.Forms.CheckBox checkBox1;

    }
}