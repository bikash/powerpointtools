namespace MiKTeXTest
{
    partial class Form1
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
            this.outputImage = new System.Windows.Forms.PictureBox();
            this.runMiKTeX = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.heightInfo = new System.Windows.Forms.TextBox();
            this.depthInfo = new System.Windows.Forms.TextBox();
            this.codeBox = new System.Windows.Forms.TextBox();
            this.latexOutputBox = new System.Windows.Forms.TextBox();
            this.dvipngOutputBox = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.outputImage)).BeginInit();
            this.SuspendLayout();
            // 
            // outputImage
            // 
            this.outputImage.Location = new System.Drawing.Point(12, 12);
            this.outputImage.Name = "outputImage";
            this.outputImage.Size = new System.Drawing.Size(127, 114);
            this.outputImage.TabIndex = 0;
            this.outputImage.TabStop = false;
            // 
            // runMiKTeX
            // 
            this.runMiKTeX.Location = new System.Drawing.Point(12, 132);
            this.runMiKTeX.Name = "runMiKTeX";
            this.runMiKTeX.Size = new System.Drawing.Size(127, 23);
            this.runMiKTeX.TabIndex = 1;
            this.runMiKTeX.Text = "Run MiKTeX";
            this.runMiKTeX.UseVisualStyleBackColor = true;
            this.runMiKTeX.Click += new System.EventHandler(this.runMiKTeX_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(24, 189);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(41, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Height:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(26, 209);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(39, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Depth:";
            // 
            // heightInfo
            // 
            this.heightInfo.Location = new System.Drawing.Point(71, 186);
            this.heightInfo.Name = "heightInfo";
            this.heightInfo.ReadOnly = true;
            this.heightInfo.Size = new System.Drawing.Size(68, 20);
            this.heightInfo.TabIndex = 4;
            // 
            // depthInfo
            // 
            this.depthInfo.Location = new System.Drawing.Point(71, 206);
            this.depthInfo.Name = "depthInfo";
            this.depthInfo.ReadOnly = true;
            this.depthInfo.ShortcutsEnabled = false;
            this.depthInfo.Size = new System.Drawing.Size(68, 20);
            this.depthInfo.TabIndex = 5;
            // 
            // codeBox
            // 
            this.codeBox.Location = new System.Drawing.Point(145, 12);
            this.codeBox.Multiline = true;
            this.codeBox.Name = "codeBox";
            this.codeBox.Size = new System.Drawing.Size(228, 212);
            this.codeBox.TabIndex = 6;
            // 
            // latexOutputBox
            // 
            this.latexOutputBox.Location = new System.Drawing.Point(12, 232);
            this.latexOutputBox.Multiline = true;
            this.latexOutputBox.Name = "latexOutputBox";
            this.latexOutputBox.ReadOnly = true;
            this.latexOutputBox.ShortcutsEnabled = false;
            this.latexOutputBox.Size = new System.Drawing.Size(361, 85);
            this.latexOutputBox.TabIndex = 7;
            // 
            // dvipngOutputBox
            // 
            this.dvipngOutputBox.Location = new System.Drawing.Point(12, 328);
            this.dvipngOutputBox.Multiline = true;
            this.dvipngOutputBox.Name = "dvipngOutputBox";
            this.dvipngOutputBox.ReadOnly = true;
            this.dvipngOutputBox.ShortcutsEnabled = false;
            this.dvipngOutputBox.Size = new System.Drawing.Size(361, 85);
            this.dvipngOutputBox.TabIndex = 8;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(385, 425);
            this.Controls.Add(this.dvipngOutputBox);
            this.Controls.Add(this.latexOutputBox);
            this.Controls.Add(this.codeBox);
            this.Controls.Add(this.depthInfo);
            this.Controls.Add(this.heightInfo);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.runMiKTeX);
            this.Controls.Add(this.outputImage);
            this.Name = "Form1";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.outputImage)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox outputImage;
        private System.Windows.Forms.Button runMiKTeX;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox heightInfo;
        private System.Windows.Forms.TextBox depthInfo;
        private System.Windows.Forms.TextBox codeBox;
        private System.Windows.Forms.TextBox latexOutputBox;
        private System.Windows.Forms.TextBox dvipngOutputBox;
    }
}

