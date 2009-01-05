namespace PowerPointLazySlides
{
    partial class LazySlidesRibbon
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
            Microsoft.Office.Tools.Ribbon.RibbonGroup lazySlideGroup;
            this.createLazySlideButton = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.separator1 = new Microsoft.Office.Tools.Ribbon.RibbonSeparator();
            this.makeNormalSlide = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.tab1 = new Microsoft.Office.Tools.Ribbon.RibbonTab();
            this.lazyShape = new Microsoft.Office.Tools.Ribbon.RibbonGroup();
            this.IncludeInheritedTextToggle = new Microsoft.Office.Tools.Ribbon.RibbonToggleButton();
            lazySlideGroup = new Microsoft.Office.Tools.Ribbon.RibbonGroup();
            lazySlideGroup.SuspendLayout();
            this.tab1.SuspendLayout();
            this.lazyShape.SuspendLayout();
            this.SuspendLayout();
            // 
            // lazySlideGroup
            // 
            lazySlideGroup.Items.Add(this.createLazySlideButton);
            lazySlideGroup.Items.Add(this.separator1);
            lazySlideGroup.Items.Add(this.makeNormalSlide);
            lazySlideGroup.Label = "Lazy Slides";
            lazySlideGroup.Name = "lazySlideGroup";
            // 
            // createLazySlideButton
            // 
            this.createLazySlideButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.createLazySlideButton.Label = "Create Lazy Slide";
            this.createLazySlideButton.Name = "createLazySlideButton";
            this.createLazySlideButton.ShowImage = true;
            this.createLazySlideButton.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.createLazySlideButton_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // makeNormalSlide
            // 
            this.makeNormalSlide.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.makeNormalSlide.Label = "Make Normal Slide";
            this.makeNormalSlide.Name = "makeNormalSlide";
            this.makeNormalSlide.ShowImage = true;
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(lazySlideGroup);
            this.tab1.Groups.Add(this.lazyShape);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // lazyShape
            // 
            this.lazyShape.Items.Add(this.IncludeInheritedTextToggle);
            this.lazyShape.Label = "Lazy Shape";
            this.lazyShape.Name = "lazyShape";
            // 
            // IncludeInheritedTextToggle
            // 
            this.IncludeInheritedTextToggle.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.IncludeInheritedTextToggle.Enabled = false;
            this.IncludeInheritedTextToggle.Label = "Include Inherited Text";
            this.IncludeInheritedTextToggle.Name = "IncludeInheritedTextToggle";
            this.IncludeInheritedTextToggle.ShowImage = true;
            this.IncludeInheritedTextToggle.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.IncludeInheritedTextToggle_Click);
            // 
            // LazySlidesRibbon
            // 
            this.Name = "LazySlidesRibbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonUIEventArgs>(this.LazySlidesRibbon_Load);
            lazySlideGroup.ResumeLayout(false);
            lazySlideGroup.PerformLayout();
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.lazyShape.ResumeLayout(false);
            this.lazyShape.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton createLazySlideButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton makeNormalSlide;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup lazyShape;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton IncludeInheritedTextToggle;
    }

    partial class ThisRibbonCollection : Microsoft.Office.Tools.Ribbon.RibbonReadOnlyCollection
    {
        internal LazySlidesRibbon LazySlidesRibbon
        {
            get { return this.GetRibbon<LazySlidesRibbon>(); }
        }
    }
}
