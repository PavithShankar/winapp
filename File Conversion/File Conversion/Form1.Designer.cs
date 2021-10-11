namespace File_Conversion
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
            this.components = new System.ComponentModel.Container();
            this.OpenFileNames = new System.Windows.Forms.ListBox();
            this.FormatFileNames = new System.Windows.Forms.ListBox();
            this.Open = new System.Windows.Forms.Button();
            this.DramaToExcel = new System.Windows.Forms.Button();
            this.MoviesToExcel = new System.Windows.Forms.Button();
            this.ProgBar = new System.Windows.Forms.ProgressBar();
            this.ExcelToJson = new System.Windows.Forms.Button();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.SuspendLayout();
            // 
            // OpenFileNames
            // 
            this.OpenFileNames.FormattingEnabled = true;
            this.OpenFileNames.HorizontalScrollbar = true;
            this.OpenFileNames.Location = new System.Drawing.Point(26, 8);
            this.OpenFileNames.Name = "OpenFileNames";
            this.OpenFileNames.Size = new System.Drawing.Size(354, 134);
            this.OpenFileNames.TabIndex = 0;
            // 
            // FormatFileNames
            // 
            this.FormatFileNames.FormattingEnabled = true;
            this.FormatFileNames.HorizontalScrollbar = true;
            this.FormatFileNames.Location = new System.Drawing.Point(26, 281);
            this.FormatFileNames.Name = "FormatFileNames";
            this.FormatFileNames.Size = new System.Drawing.Size(354, 134);
            this.FormatFileNames.TabIndex = 1;
            // 
            // Open
            // 
            this.Open.Location = new System.Drawing.Point(160, 149);
            this.Open.Name = "Open";
            this.Open.Size = new System.Drawing.Size(75, 23);
            this.Open.TabIndex = 2;
            this.Open.Text = "Open";
            this.Open.UseVisualStyleBackColor = true;
            this.Open.Click += new System.EventHandler(this.Open_Click);
            // 
            // DramaToExcel
            // 
            this.DramaToExcel.Location = new System.Drawing.Point(627, 103);
            this.DramaToExcel.Name = "DramaToExcel";
            this.DramaToExcel.Size = new System.Drawing.Size(130, 39);
            this.DramaToExcel.TabIndex = 3;
            this.DramaToExcel.Text = "Drama To Excel";
            this.DramaToExcel.UseVisualStyleBackColor = true;
            this.DramaToExcel.Click += new System.EventHandler(this.DramaToExcel_Click);
            // 
            // MoviesToExcel
            // 
            this.MoviesToExcel.Location = new System.Drawing.Point(627, 213);
            this.MoviesToExcel.Name = "MoviesToExcel";
            this.MoviesToExcel.Size = new System.Drawing.Size(130, 39);
            this.MoviesToExcel.TabIndex = 4;
            this.MoviesToExcel.Text = "Movies to Excel";
            this.MoviesToExcel.UseVisualStyleBackColor = true;
            this.MoviesToExcel.Click += new System.EventHandler(this.MoviesToExcel_Click);
            // 
            // ProgBar
            // 
            this.ProgBar.Location = new System.Drawing.Point(133, 460);
            this.ProgBar.Name = "ProgBar";
            this.ProgBar.Size = new System.Drawing.Size(598, 23);
            this.ProgBar.TabIndex = 5;
            // 
            // ExcelToJson
            // 
            this.ExcelToJson.Location = new System.Drawing.Point(627, 329);
            this.ExcelToJson.Name = "ExcelToJson";
            this.ExcelToJson.Size = new System.Drawing.Size(130, 41);
            this.ExcelToJson.TabIndex = 6;
            this.ExcelToJson.Text = "Excel To Json";
            this.ExcelToJson.UseVisualStyleBackColor = true;
            this.ExcelToJson.Click += new System.EventHandler(this.ExcelToJson_Click);
            // 
            // timer1
            // 
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(864, 530);
            this.Controls.Add(this.ExcelToJson);
            this.Controls.Add(this.ProgBar);
            this.Controls.Add(this.MoviesToExcel);
            this.Controls.Add(this.DramaToExcel);
            this.Controls.Add(this.Open);
            this.Controls.Add(this.FormatFileNames);
            this.Controls.Add(this.OpenFileNames);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ListBox OpenFileNames;
        private System.Windows.Forms.ListBox FormatFileNames;
        private System.Windows.Forms.Button Open;
        private System.Windows.Forms.Button DramaToExcel;
        private System.Windows.Forms.Button MoviesToExcel;
        private System.Windows.Forms.ProgressBar ProgBar;
        private System.Windows.Forms.Button ExcelToJson;
        private System.Windows.Forms.Timer timer1;
    }
}

