
namespace VKSMM
{
    partial class EnterPassForm
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
            this.buttonTestPass = new System.Windows.Forms.Button();
            this.buttonIgnorePass = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.buttonTestPass);
            this.groupBox1.Controls.Add(this.buttonIgnorePass);
            this.groupBox1.Controls.Add(this.textBox1);
            this.groupBox1.Location = new System.Drawing.Point(12, 8);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(330, 113);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Введите пароль АДМИНИСТРАТОРА";
            // 
            // buttonTestPass
            // 
            this.buttonTestPass.Location = new System.Drawing.Point(231, 74);
            this.buttonTestPass.Name = "buttonTestPass";
            this.buttonTestPass.Size = new System.Drawing.Size(75, 23);
            this.buttonTestPass.TabIndex = 1;
            this.buttonTestPass.Text = "ПРИНЯТЬ";
            this.buttonTestPass.UseVisualStyleBackColor = true;
            this.buttonTestPass.Click += new System.EventHandler(this.buttonTestPass_Click);
            // 
            // buttonIgnorePass
            // 
            this.buttonIgnorePass.Location = new System.Drawing.Point(22, 74);
            this.buttonIgnorePass.Name = "buttonIgnorePass";
            this.buttonIgnorePass.Size = new System.Drawing.Size(75, 23);
            this.buttonIgnorePass.TabIndex = 1;
            this.buttonIgnorePass.Text = "ОТМЕНА";
            this.buttonIgnorePass.UseVisualStyleBackColor = true;
            this.buttonIgnorePass.Click += new System.EventHandler(this.buttonIgnorePass_Click);
            // 
            // textBox1
            // 
            this.textBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBox1.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.textBox1.Location = new System.Drawing.Point(22, 30);
            this.textBox1.Name = "textBox1";
            this.textBox1.PasswordChar = '*';
            this.textBox1.Size = new System.Drawing.Size(284, 20);
            this.textBox1.TabIndex = 0;
            // 
            // EnterPassForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(355, 134);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "EnterPassForm";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "EnterPassForm";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button buttonTestPass;
        private System.Windows.Forms.Button buttonIgnorePass;
        private System.Windows.Forms.TextBox textBox1;
    }
}