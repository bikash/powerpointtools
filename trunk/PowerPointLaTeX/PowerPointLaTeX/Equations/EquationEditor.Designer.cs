namespace PowerPointLaTeX {
    partial class EquationEditor {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            System.Windows.Forms.Button applyButton;
            System.Windows.Forms.Button cancelButton;
            System.Windows.Forms.Label label1;
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.flowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            this.formulaText = new System.Windows.Forms.TextBox();
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.formulaPreview = new System.Windows.Forms.PictureBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.fontSizeUpDown = new System.Windows.Forms.NumericUpDown();
            applyButton = new System.Windows.Forms.Button();
            cancelButton = new System.Windows.Forms.Button();
            label1 = new System.Windows.Forms.Label();
            this.tableLayoutPanel1.SuspendLayout();
            this.flowLayoutPanel1.SuspendLayout();
            this.tableLayoutPanel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.formulaPreview)).BeginInit();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.fontSizeUpDown)).BeginInit();
            this.SuspendLayout();
            // 
            // applyButton
            // 
            applyButton.AutoSize = true;
            applyButton.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            applyButton.DialogResult = System.Windows.Forms.DialogResult.OK;
            applyButton.Location = new System.Drawing.Point(82, 3);
            applyButton.Name = "applyButton";
            applyButton.Size = new System.Drawing.Size(88, 23);
            applyButton.TabIndex = 0;
            applyButton.Text = "Apply Changes";
            applyButton.UseVisualStyleBackColor = true;
            applyButton.Click += new System.EventHandler(this.applyButton_Click);
            // 
            // cancelButton
            // 
            cancelButton.AutoSize = true;
            cancelButton.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            cancelButton.Location = new System.Drawing.Point(176, 3);
            cancelButton.Name = "cancelButton";
            cancelButton.Size = new System.Drawing.Size(50, 23);
            cancelButton.TabIndex = 1;
            cancelButton.Text = "Cancel";
            cancelButton.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new System.Drawing.Point(3, 8);
            label1.Name = "label1";
            label1.Size = new System.Drawing.Size(54, 13);
            label1.TabIndex = 6;
            label1.Text = "Font Size:";
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 235F));
            this.tableLayoutPanel1.Controls.Add(this.flowLayoutPanel1, 1, 2);
            this.tableLayoutPanel1.Controls.Add(this.formulaText, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.tableLayoutPanel2, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.panel1, 0, 2);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 3;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 102F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 8F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(459, 250);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // flowLayoutPanel1
            // 
            this.flowLayoutPanel1.AutoSize = true;
            this.flowLayoutPanel1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.flowLayoutPanel1.Controls.Add(cancelButton);
            this.flowLayoutPanel1.Controls.Add(applyButton);
            this.flowLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.flowLayoutPanel1.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft;
            this.flowLayoutPanel1.Location = new System.Drawing.Point(227, 218);
            this.flowLayoutPanel1.Name = "flowLayoutPanel1";
            this.flowLayoutPanel1.Size = new System.Drawing.Size(229, 29);
            this.flowLayoutPanel1.TabIndex = 1;
            // 
            // formulaText
            // 
            this.formulaText.AcceptsReturn = true;
            this.formulaText.AcceptsTab = true;
            this.tableLayoutPanel1.SetColumnSpan(this.formulaText, 2);
            this.formulaText.Dock = System.Windows.Forms.DockStyle.Fill;
            this.formulaText.Location = new System.Drawing.Point(3, 116);
            this.formulaText.Multiline = true;
            this.formulaText.Name = "formulaText";
            this.formulaText.Size = new System.Drawing.Size(453, 96);
            this.formulaText.TabIndex = 2;
            this.formulaText.UseSystemPasswordChar = true;
            this.formulaText.TextChanged += new System.EventHandler(this.formulaText_TextChanged);
            // 
            // tableLayoutPanel2
            // 
            this.tableLayoutPanel2.AutoScroll = true;
            this.tableLayoutPanel2.AutoSize = true;
            this.tableLayoutPanel2.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.tableLayoutPanel2.ColumnCount = 1;
            this.tableLayoutPanel1.SetColumnSpan(this.tableLayoutPanel2, 2);
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel2.Controls.Add(this.formulaPreview, 0, 0);
            this.tableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel2.GrowStyle = System.Windows.Forms.TableLayoutPanelGrowStyle.AddColumns;
            this.tableLayoutPanel2.Location = new System.Drawing.Point(3, 3);
            this.tableLayoutPanel2.Name = "tableLayoutPanel2";
            this.tableLayoutPanel2.RowCount = 1;
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel2.Size = new System.Drawing.Size(453, 107);
            this.tableLayoutPanel2.TabIndex = 3;
            // 
            // formulaPreview
            // 
            this.formulaPreview.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.formulaPreview.Cursor = System.Windows.Forms.Cursors.Arrow;
            this.formulaPreview.Location = new System.Drawing.Point(221, 48);
            this.formulaPreview.Name = "formulaPreview";
            this.formulaPreview.Size = new System.Drawing.Size(10, 10);
            this.formulaPreview.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.formulaPreview.TabIndex = 6;
            this.formulaPreview.TabStop = false;
            // 
            // panel1
            // 
            this.panel1.AutoSize = true;
            this.panel1.Controls.Add(this.fontSizeUpDown);
            this.panel1.Controls.Add(label1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(3, 218);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(218, 29);
            this.panel1.TabIndex = 4;
            // 
            // fontSizeUpDown
            // 
            this.fontSizeUpDown.Location = new System.Drawing.Point(63, 6);
            this.fontSizeUpDown.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.fontSizeUpDown.Name = "fontSizeUpDown";
            this.fontSizeUpDown.Size = new System.Drawing.Size(42, 20);
            this.fontSizeUpDown.TabIndex = 7;
            this.fontSizeUpDown.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.fontSizeUpDown.ValueChanged += new System.EventHandler(this.fontSizeUpDown_ValueChanged);
            this.fontSizeUpDown.Click += new System.EventHandler(this.fontSizeUpDown_Click);
            // 
            // EquationEditor
            // 
            this.AcceptButton = applyButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = cancelButton;
            this.ClientSize = new System.Drawing.Size(459, 250);
            this.Controls.Add(this.tableLayoutPanel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow;
            this.Name = "EquationEditor";
            this.Text = "Formula Editor";
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.flowLayoutPanel1.ResumeLayout(false);
            this.flowLayoutPanel1.PerformLayout();
            this.tableLayoutPanel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.formulaPreview)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.fontSizeUpDown)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel1;
        private System.Windows.Forms.TextBox formulaText;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel2;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.NumericUpDown fontSizeUpDown;
        private System.Windows.Forms.PictureBox formulaPreview;




    }
}