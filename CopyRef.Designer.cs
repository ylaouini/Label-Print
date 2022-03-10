namespace Label_Print
{
    partial class CopyRef
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
            this.txtNewEpn = new System.Windows.Forms.TextBox();
            this.btnDupplicate = new Guna.UI2.WinForms.Guna2Button();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // txtNewEpn
            // 
            this.txtNewEpn.Font = new System.Drawing.Font("Segoe UI", 10F);
            this.txtNewEpn.Location = new System.Drawing.Point(35, 46);
            this.txtNewEpn.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.txtNewEpn.Name = "txtNewEpn";
            this.txtNewEpn.Size = new System.Drawing.Size(252, 25);
            this.txtNewEpn.TabIndex = 0;
            // 
            // btnDupplicate
            // 
            this.btnDupplicate.Animated = true;
            this.btnDupplicate.CheckedState.Parent = this.btnDupplicate;
            this.btnDupplicate.CustomImages.Parent = this.btnDupplicate;
            this.btnDupplicate.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.btnDupplicate.ForeColor = System.Drawing.Color.White;
            this.btnDupplicate.HoverState.Parent = this.btnDupplicate;
            this.btnDupplicate.Location = new System.Drawing.Point(111, 85);
            this.btnDupplicate.Name = "btnDupplicate";
            this.btnDupplicate.ShadowDecoration.Parent = this.btnDupplicate;
            this.btnDupplicate.Size = new System.Drawing.Size(100, 33);
            this.btnDupplicate.TabIndex = 73;
            this.btnDupplicate.Text = "OK";
            this.btnDupplicate.Click += new System.EventHandler(this.btnDupplicate_Click);
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 10F);
            this.label1.Location = new System.Drawing.Point(139, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(45, 19);
            this.label1.TabIndex = 74;
            this.label1.Text = "label1";
            // 
            // CopyRef
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Window;
            this.ClientSize = new System.Drawing.Size(315, 132);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnDupplicate);
            this.Controls.Add(this.txtNewEpn);
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Name = "CopyRef";
            this.Text = "CopyRef";
            this.Load += new System.EventHandler(this.CopyRef_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtNewEpn;
        private Guna.UI2.WinForms.Guna2Button btnDupplicate;
        private System.Windows.Forms.Label label1;
    }
}