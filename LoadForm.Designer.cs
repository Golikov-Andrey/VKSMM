
namespace VKSMM
{
    partial class LoadForm
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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.progressBarLoadForm = new System.Windows.Forms.ProgressBar();
            this.labelLoadForm = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.progressBarLoadForm);
            this.groupBox1.Controls.Add(this.labelLoadForm);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(472, 126);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Подождите идет загрузка";
            // 
            // progressBarLoadForm
            // 
            this.progressBarLoadForm.Location = new System.Drawing.Point(16, 72);
            this.progressBarLoadForm.Name = "progressBarLoadForm";
            this.progressBarLoadForm.Size = new System.Drawing.Size(438, 20);
            this.progressBarLoadForm.TabIndex = 1;
            // 
            // labelLoadForm
            // 
            this.labelLoadForm.AutoSize = true;
            this.labelLoadForm.Location = new System.Drawing.Point(29, 41);
            this.labelLoadForm.Name = "labelLoadForm";
            this.labelLoadForm.Size = new System.Drawing.Size(57, 13);
            this.labelLoadForm.TabIndex = 0;
            this.labelLoadForm.Text = "Загрузка:";
            // 
            // LoadForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(497, 152);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "LoadForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "LoadForm";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        public System.Windows.Forms.ProgressBar progressBarLoadForm;
        public System.Windows.Forms.Label labelLoadForm;
    }
}