namespace LeadsExtractor
{
    partial class frmSearchEnginesEdit
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
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtSearchEngineName = new System.Windows.Forms.TextBox();
            this.txtSearchEngineURL = new System.Windows.Forms.TextBox();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(62, 124);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "Ok";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(153, 124);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 1;
            this.button2.Text = "Cancel";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(111, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Search Engine Name:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(13, 54);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(105, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Search Engine URL:";
            // 
            // txtSearchEngineName
            // 
            this.txtSearchEngineName.Location = new System.Drawing.Point(13, 30);
            this.txtSearchEngineName.Name = "txtSearchEngineName";
            this.txtSearchEngineName.Size = new System.Drawing.Size(259, 20);
            this.txtSearchEngineName.TabIndex = 5;
            this.txtSearchEngineName.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSearchEngineName_KeyDown);
            // 
            // txtSearchEngineURL
            // 
            this.txtSearchEngineURL.Location = new System.Drawing.Point(13, 71);
            this.txtSearchEngineURL.Name = "txtSearchEngineURL";
            this.txtSearchEngineURL.Size = new System.Drawing.Size(259, 20);
            this.txtSearchEngineURL.TabIndex = 6;
            this.txtSearchEngineURL.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSearchEngineURL_KeyDown);
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(13, 95);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(82, 17);
            this.checkBox1.TabIndex = 7;
            this.checkBox1.Text = "Is Enabled?";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // frmSearchEnginesEdit
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 160);
            this.Controls.Add(this.checkBox1);
            this.Controls.Add(this.txtSearchEngineURL);
            this.Controls.Add(this.txtSearchEngineName);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmSearchEnginesEdit";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Search Engine";
            this.Load += new System.EventHandler(this.frmSearchEnginesEdit_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtSearchEngineName;
        private System.Windows.Forms.TextBox txtSearchEngineURL;
        private System.Windows.Forms.CheckBox checkBox1;
    }
}