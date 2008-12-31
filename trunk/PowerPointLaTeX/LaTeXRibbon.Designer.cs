namespace PowerPointLaTeX
{
    partial class LaTeXRibbon
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

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(LaTeXRibbon));
            this.LaTeX = new Microsoft.Office.Tools.Ribbon.RibbonTab();
            this.generalGroup = new Microsoft.Office.Tools.Ribbon.RibbonGroup();
            this.button1 = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.offlineModeToggle = new Microsoft.Office.Tools.Ribbon.RibbonToggleButton();
            this.separator1 = new Microsoft.Office.Tools.Ribbon.RibbonSeparator();
            this.bakeButton = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.inlineGroup = new Microsoft.Office.Tools.Ribbon.RibbonGroup();
            this.automaticCompilationToggle = new Microsoft.Office.Tools.Ribbon.RibbonToggleButton();
            this.CompileSlide = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.DecompileSlide = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.LaTeX.SuspendLayout();
            this.generalGroup.SuspendLayout();
            this.inlineGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // LaTeX
            // 
            this.LaTeX.Groups.Add(this.generalGroup);
            this.LaTeX.Groups.Add(this.inlineGroup);
            this.LaTeX.Label = "LaTeX";
            this.LaTeX.Name = "LaTeX";
            // 
            // generalGroup
            // 
            this.generalGroup.Items.Add(this.button1);
            this.generalGroup.Items.Add(this.offlineModeToggle);
            this.generalGroup.Items.Add(this.separator1);
            this.generalGroup.Items.Add(this.bakeButton);
            this.generalGroup.Label = "General";
            this.generalGroup.Name = "generalGroup";
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Label = "Preferences";
            this.button1.Name = "button1";
            this.button1.OfficeImageId = "MessageOptions";
            this.button1.ShowImage = true;
            // 
            // offlineModeToggle
            // 
            this.offlineModeToggle.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.offlineModeToggle.Description = "Compiles everything and protects it from changes";
            this.offlineModeToggle.Image = ((System.Drawing.Image) (resources.GetObject("offlineModeToggle.Image")));
            this.offlineModeToggle.Label = "Offline Mode";
            this.offlineModeToggle.Name = "offlineModeToggle";
            this.offlineModeToggle.ShowImage = true;
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // bakeButton
            // 
            this.bakeButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.bakeButton.Description = "Compile everything and remove all meta-information";
            this.bakeButton.Label = "Bake LaTeX";
            this.bakeButton.Name = "bakeButton";
            this.bakeButton.OfficeImageId = "FileMarkAsFinal";
            this.bakeButton.ShowImage = true;
            // 
            // inlineGroup
            // 
            this.inlineGroup.Items.Add(this.automaticCompilationToggle);
            this.inlineGroup.Items.Add(this.CompileSlide);
            this.inlineGroup.Items.Add(this.DecompileSlide);
            this.inlineGroup.Label = "Inline Formulas";
            this.inlineGroup.Name = "inlineGroup";
            // 
            // automaticCompilationToggle
            // 
            this.automaticCompilationToggle.Checked = true;
            this.automaticCompilationToggle.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.automaticCompilationToggle.Description = "Automatically compile LaTeX shortcodes";
            this.automaticCompilationToggle.Image = ((System.Drawing.Image) (resources.GetObject("automaticCompilationToggle.Image")));
            this.automaticCompilationToggle.Label = "Automatic Compilation";
            this.automaticCompilationToggle.Name = "automaticCompilationToggle";
            this.automaticCompilationToggle.ShowImage = true;
            this.automaticCompilationToggle.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.automaticCompilationToggle_Click);
            // 
            // CompileSlide
            // 
            this.CompileSlide.Description = "Compile the current selection/slide";
            this.CompileSlide.Label = "Compile Selection";
            this.CompileSlide.Name = "CompileSlide";
            this.CompileSlide.OfficeImageId = "MacroPlay";
            this.CompileSlide.ShowImage = true;
            this.CompileSlide.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.button1_Click);
            // 
            // DecompileSlide
            // 
            this.DecompileSlide.Label = "Decompile Selection";
            this.DecompileSlide.Name = "DecompileSlide";
            this.DecompileSlide.OfficeImageId = "MailMergeGoToPreviousRecord";
            this.DecompileSlide.ShowImage = true;
            this.DecompileSlide.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.DecompileSlide_Click);
            // 
            // LaTeXRibbon
            // 
            this.Name = "LaTeXRibbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.LaTeX);
            this.Load += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonUIEventArgs>(this.LaTeXRibbon_Load);
            this.LaTeX.ResumeLayout(false);
            this.LaTeX.PerformLayout();
            this.generalGroup.ResumeLayout(false);
            this.generalGroup.PerformLayout();
            this.inlineGroup.ResumeLayout(false);
            this.inlineGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab LaTeX;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup inlineGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CompileSlide;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton DecompileSlide;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup generalGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton offlineModeToggle;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton automaticCompilationToggle;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bakeButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
    }

    partial class ThisRibbonCollection : Microsoft.Office.Tools.Ribbon.RibbonReadOnlyCollection
    {
        internal LaTeXRibbon LaTeXRibbon
        {
            get { return this.GetRibbon<LaTeXRibbon>(); }
        }
    }
}
