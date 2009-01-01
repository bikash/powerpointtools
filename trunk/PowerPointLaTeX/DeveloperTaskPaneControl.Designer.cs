namespace PowerPointLaTeX
{
    partial class DeveloperTaskPaneControl
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
            this.tabControl = new System.Windows.Forms.TabControl();
            this.TagsPage = new System.Windows.Forms.TabPage();
            this.tagsLayout = new System.Windows.Forms.FlowLayoutPanel();
            this.GeneralPage = new System.Windows.Forms.TabPage();
            this.tabControl.SuspendLayout();
            this.TagsPage.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl
            // 
            this.tabControl.Controls.Add(this.TagsPage);
            this.tabControl.Controls.Add(this.GeneralPage);
            this.tabControl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl.Location = new System.Drawing.Point(0, 0);
            this.tabControl.Name = "tabControl";
            this.tabControl.SelectedIndex = 0;
            this.tabControl.Size = new System.Drawing.Size(400, 491);
            this.tabControl.TabIndex = 0;
            // 
            // TagsPage
            // 
            this.TagsPage.Controls.Add(this.tagsLayout);
            this.TagsPage.Location = new System.Drawing.Point(4, 22);
            this.TagsPage.Name = "TagsPage";
            this.TagsPage.Padding = new System.Windows.Forms.Padding(3);
            this.TagsPage.Size = new System.Drawing.Size(392, 465);
            this.TagsPage.TabIndex = 0;
            this.TagsPage.Text = "Tags";
            this.TagsPage.UseVisualStyleBackColor = true;
            // 
            // tagsLayout
            // 
            this.tagsLayout.AutoScroll = true;
            this.tagsLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tagsLayout.Location = new System.Drawing.Point(3, 3);
            this.tagsLayout.Name = "tagsLayout";
            this.tagsLayout.Size = new System.Drawing.Size(386, 459);
            this.tagsLayout.TabIndex = 3;
            // 
            // GeneralPage
            // 
            this.GeneralPage.Location = new System.Drawing.Point(4, 22);
            this.GeneralPage.Name = "GeneralPage";
            this.GeneralPage.Padding = new System.Windows.Forms.Padding(3);
            this.GeneralPage.Size = new System.Drawing.Size(392, 465);
            this.GeneralPage.TabIndex = 1;
            this.GeneralPage.Text = "General";
            this.GeneralPage.UseVisualStyleBackColor = true;
            // 
            // DeveloperTaskPaneControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.tabControl);
            this.Name = "DeveloperTaskPaneControl";
            this.Size = new System.Drawing.Size(400, 491);
            this.tabControl.ResumeLayout(false);
            this.TagsPage.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl;
        private System.Windows.Forms.TabPage TagsPage;
        private System.Windows.Forms.TabPage GeneralPage;
        private System.Windows.Forms.FlowLayoutPanel tagsLayout;
    }
}
