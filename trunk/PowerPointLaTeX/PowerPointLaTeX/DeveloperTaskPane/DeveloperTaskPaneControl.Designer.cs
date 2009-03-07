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
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.tagsLayout = new System.Windows.Forms.FlowLayoutPanel();
            this.flowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            this.refreshButton = new System.Windows.Forms.Button();
            this.selectAll = new System.Windows.Forms.Button();
            this.useCurrentSelectionButton = new System.Windows.Forms.Button();
            this.GeneralPage = new System.Windows.Forms.TabPage();
            this.tabControl.SuspendLayout();
            this.TagsPage.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            this.flowLayoutPanel1.SuspendLayout();
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
            this.tabControl.Size = new System.Drawing.Size(319, 424);
            this.tabControl.TabIndex = 0;
            // 
            // TagsPage
            // 
            this.TagsPage.Controls.Add(this.tableLayoutPanel1);
            this.TagsPage.Location = new System.Drawing.Point(4, 22);
            this.TagsPage.Name = "TagsPage";
            this.TagsPage.Padding = new System.Windows.Forms.Padding(3);
            this.TagsPage.Size = new System.Drawing.Size(311, 398);
            this.TagsPage.TabIndex = 0;
            this.TagsPage.Text = "Tags";
            this.TagsPage.UseVisualStyleBackColor = true;
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Controls.Add(this.tagsLayout, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.flowLayoutPanel1, 0, 0);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(3, 3);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 2;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.Size = new System.Drawing.Size(305, 392);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // tagsLayout
            // 
            this.tagsLayout.AutoScroll = true;
            this.tagsLayout.AutoSize = true;
            this.tagsLayout.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.tagsLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tagsLayout.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.tagsLayout.Location = new System.Drawing.Point(3, 38);
            this.tagsLayout.Name = "tagsLayout";
            this.tagsLayout.Size = new System.Drawing.Size(299, 351);
            this.tagsLayout.TabIndex = 4;
            // 
            // flowLayoutPanel1
            // 
            this.flowLayoutPanel1.AutoSize = true;
            this.flowLayoutPanel1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.flowLayoutPanel1.Controls.Add(this.refreshButton);
            this.flowLayoutPanel1.Controls.Add(this.selectAll);
            this.flowLayoutPanel1.Controls.Add(this.useCurrentSelectionButton);
            this.flowLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.flowLayoutPanel1.Location = new System.Drawing.Point(3, 3);
            this.flowLayoutPanel1.Name = "flowLayoutPanel1";
            this.flowLayoutPanel1.Size = new System.Drawing.Size(299, 29);
            this.flowLayoutPanel1.TabIndex = 5;
            // 
            // refreshButton
            // 
            this.refreshButton.AutoSize = true;
            this.refreshButton.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.refreshButton.Location = new System.Drawing.Point(3, 3);
            this.refreshButton.Name = "refreshButton";
            this.refreshButton.Size = new System.Drawing.Size(54, 23);
            this.refreshButton.TabIndex = 1;
            this.refreshButton.Text = "Refresh";
            this.refreshButton.UseVisualStyleBackColor = true;
            this.refreshButton.Click += new System.EventHandler(this.refreshButton_Click);
            // 
            // selectAll
            // 
            this.selectAll.AutoSize = true;
            this.selectAll.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.selectAll.Dock = System.Windows.Forms.DockStyle.Fill;
            this.selectAll.Location = new System.Drawing.Point(63, 3);
            this.selectAll.Name = "selectAll";
            this.selectAll.Size = new System.Drawing.Size(61, 23);
            this.selectAll.TabIndex = 2;
            this.selectAll.Text = "Select All";
            this.selectAll.UseVisualStyleBackColor = true;
            this.selectAll.Click += new System.EventHandler(this.selectAllButton_Click);
            // 
            // useCurrentSelectionButton
            // 
            this.useCurrentSelectionButton.AutoSize = true;
            this.useCurrentSelectionButton.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.useCurrentSelectionButton.Dock = System.Windows.Forms.DockStyle.Fill;
            this.useCurrentSelectionButton.Location = new System.Drawing.Point(130, 3);
            this.useCurrentSelectionButton.Name = "useCurrentSelectionButton";
            this.useCurrentSelectionButton.Size = new System.Drawing.Size(120, 23);
            this.useCurrentSelectionButton.TabIndex = 3;
            this.useCurrentSelectionButton.Text = "Use Current Selection";
            this.useCurrentSelectionButton.UseVisualStyleBackColor = true;
            this.useCurrentSelectionButton.Click += new System.EventHandler(this.useCurrentSelectionButton_Click);
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
            this.Size = new System.Drawing.Size(319, 424);
            this.tabControl.ResumeLayout(false);
            this.TagsPage.ResumeLayout(false);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.flowLayoutPanel1.ResumeLayout(false);
            this.flowLayoutPanel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl;
        private System.Windows.Forms.TabPage TagsPage;
        private System.Windows.Forms.TabPage GeneralPage;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.FlowLayoutPanel tagsLayout;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel1;
        private System.Windows.Forms.Button refreshButton;
        private System.Windows.Forms.Button selectAll;
        private System.Windows.Forms.Button useCurrentSelectionButton;
    }
}
