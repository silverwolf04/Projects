namespace OracleTest
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
            this.btnOracleCheck = new System.Windows.Forms.Button();
            this.lbl1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnOracleCheck
            // 
            this.btnOracleCheck.Location = new System.Drawing.Point(288, 71);
            this.btnOracleCheck.Name = "btnOracleCheck";
            this.btnOracleCheck.Size = new System.Drawing.Size(75, 23);
            this.btnOracleCheck.TabIndex = 0;
            this.btnOracleCheck.Text = "Test Oracle";
            this.btnOracleCheck.UseVisualStyleBackColor = true;
            this.btnOracleCheck.Click += new System.EventHandler(this.BtnOracleCheck_Click);
            // 
            // lbl1
            // 
            this.lbl1.AutoSize = true;
            this.lbl1.Location = new System.Drawing.Point(285, 119);
            this.lbl1.Name = "lbl1";
            this.lbl1.Size = new System.Drawing.Size(61, 13);
            this.lbl1.TabIndex = 2;
            this.lbl1.Text = "Not Started";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.lbl1);
            this.Controls.Add(this.btnOracleCheck);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnOracleCheck;
        private System.Windows.Forms.Label lbl1;
    }
}

