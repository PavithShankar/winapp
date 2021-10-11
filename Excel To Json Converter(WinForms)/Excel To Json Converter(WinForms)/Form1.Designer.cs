namespace Excel_To_Json_Converter_WinForms_
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
            this.Convert = new System.Windows.Forms.Button();
            this.OutputListBox = new System.Windows.Forms.ListBox();
            this.InputListBox = new System.Windows.Forms.ListBox();
            this.Openbut = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // Convert
            // 
            this.Convert.Location = new System.Drawing.Point(569, 179);
            this.Convert.Name = "Convert";
            this.Convert.Size = new System.Drawing.Size(89, 36);
            this.Convert.TabIndex = 0;
            this.Convert.Text = "Excel to Json";
            this.Convert.UseVisualStyleBackColor = true;
            this.Convert.Click += new System.EventHandler(this.Convert_Click);
            // 
            // OutputListBox
            // 
            this.OutputListBox.FormattingEnabled = true;
            this.OutputListBox.Location = new System.Drawing.Point(45, 229);
            this.OutputListBox.Name = "OutputListBox";
            this.OutputListBox.Size = new System.Drawing.Size(401, 95);
            this.OutputListBox.TabIndex = 3;
            // 
            // InputListBox
            // 
            this.InputListBox.FormattingEnabled = true;
            this.InputListBox.Location = new System.Drawing.Point(45, 66);
            this.InputListBox.Name = "InputListBox";
            this.InputListBox.Size = new System.Drawing.Size(401, 95);
            this.InputListBox.TabIndex = 2;
            // 
            // Openbut
            // 
            this.Openbut.Location = new System.Drawing.Point(198, 167);
            this.Openbut.Name = "Openbut";
            this.Openbut.Size = new System.Drawing.Size(75, 23);
            this.Openbut.TabIndex = 4;
            this.Openbut.Text = "Open";
            this.Openbut.UseVisualStyleBackColor = true;
            this.Openbut.Click += new System.EventHandler(this.Openbut_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(817, 407);
            this.Controls.Add(this.Openbut);
            this.Controls.Add(this.OutputListBox);
            this.Controls.Add(this.InputListBox);
            this.Controls.Add(this.Convert);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button Convert;
        private System.Windows.Forms.ListBox OutputListBox;
        private System.Windows.Forms.ListBox InputListBox;
        private System.Windows.Forms.Button Openbut;
    }
}

