namespace ExcelProviderForms
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
            this.button1 = new System.Windows.Forms.Button();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.btnProcesar = new System.Windows.Forms.Button();
            this.listBox2 = new System.Windows.Forms.ListBox();
            this.labelOld = new System.Windows.Forms.Label();
            this.labelNew = new System.Windows.Forms.Label();
            this.txtOldPath = new System.Windows.Forms.TextBox();
            this.txtNewPath = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(30, 123);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(145, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "Seleccionar Carpeta";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // listBox1
            // 
            this.listBox1.FormattingEnabled = true;
            this.listBox1.Location = new System.Drawing.Point(30, 152);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(309, 368);
            this.listBox1.TabIndex = 1;
            // 
            // btnProcesar
            // 
            this.btnProcesar.Location = new System.Drawing.Point(365, 123);
            this.btnProcesar.Name = "btnProcesar";
            this.btnProcesar.Size = new System.Drawing.Size(75, 23);
            this.btnProcesar.TabIndex = 2;
            this.btnProcesar.Text = "Procesar";
            this.btnProcesar.UseVisualStyleBackColor = true;
            this.btnProcesar.Click += new System.EventHandler(this.btnProcesar_Click);
            // 
            // listBox2
            // 
            this.listBox2.FormattingEnabled = true;
            this.listBox2.Location = new System.Drawing.Point(365, 152);
            this.listBox2.Name = "listBox2";
            this.listBox2.Size = new System.Drawing.Size(423, 368);
            this.listBox2.TabIndex = 3;
            // 
            // labelOld
            // 
            this.labelOld.AutoSize = true;
            this.labelOld.Location = new System.Drawing.Point(24, 25);
            this.labelOld.Name = "labelOld";
            this.labelOld.Size = new System.Drawing.Size(48, 13);
            this.labelOld.TabIndex = 4;
            this.labelOld.Text = "Old Path";
            // 
            // labelNew
            // 
            this.labelNew.AutoSize = true;
            this.labelNew.Location = new System.Drawing.Point(24, 52);
            this.labelNew.Name = "labelNew";
            this.labelNew.Size = new System.Drawing.Size(54, 13);
            this.labelNew.TabIndex = 5;
            this.labelNew.Text = "New Path";
            // 
            // txtOldPath
            // 
            this.txtOldPath.Location = new System.Drawing.Point(96, 25);
            this.txtOldPath.Name = "txtOldPath";
            this.txtOldPath.Size = new System.Drawing.Size(240, 20);
            this.txtOldPath.TabIndex = 6;
            // 
            // txtNewPath
            // 
            this.txtNewPath.Location = new System.Drawing.Point(96, 49);
            this.txtNewPath.Name = "txtNewPath";
            this.txtNewPath.Size = new System.Drawing.Size(240, 20);
            this.txtNewPath.TabIndex = 7;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 532);
            this.Controls.Add(this.txtNewPath);
            this.Controls.Add(this.txtOldPath);
            this.Controls.Add(this.labelNew);
            this.Controls.Add(this.labelOld);
            this.Controls.Add(this.listBox2);
            this.Controls.Add(this.btnProcesar);
            this.Controls.Add(this.listBox1);
            this.Controls.Add(this.button1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.ListBox listBox1;
        private System.Windows.Forms.Button btnProcesar;
        private System.Windows.Forms.ListBox listBox2;
        private System.Windows.Forms.Label labelOld;
        private System.Windows.Forms.Label labelNew;
        private System.Windows.Forms.TextBox txtOldPath;
        private System.Windows.Forms.TextBox txtNewPath;
    }
}

