
namespace XL_jobbet
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.button1 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.button8 = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.ZnS_M18 = new System.Windows.Forms.CheckBox();
            this.ZnS_M21 = new System.Windows.Forms.CheckBox();
            this.ZnS_M20 = new System.Windows.Forms.CheckBox();
            this.ZnS_M27 = new System.Windows.Forms.CheckBox();
            this.ZnS_M19 = new System.Windows.Forms.CheckBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.ZnS_II_M18 = new System.Windows.Forms.CheckBox();
            this.ZnS_II_M21 = new System.Windows.Forms.CheckBox();
            this.ZnS_II_M20 = new System.Windows.Forms.CheckBox();
            this.ZnS_II_M27 = new System.Windows.Forms.CheckBox();
            this.ZnS_II_M19 = new System.Windows.Forms.CheckBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.ZnS_III_M18 = new System.Windows.Forms.CheckBox();
            this.ZnS_III_M21 = new System.Windows.Forms.CheckBox();
            this.ZnS_III_M20 = new System.Windows.Forms.CheckBox();
            this.ZnS_III_M27 = new System.Windows.Forms.CheckBox();
            this.ZnS_III_M19 = new System.Windows.Forms.CheckBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.Export_DT_To_Excel = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(130, 1073);
            this.button1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(320, 37);
            this.button1.TabIndex = 0;
            this.button1.Text = "Öppna_Mall-skapa_öppningar+kälpåfyll";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(130, 1127);
            this.button4.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(320, 37);
            this.button4.TabIndex = 3;
            this.button4.Text = "ExcelSaveAsNewFile";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button8
            // 
            this.button8.Location = new System.Drawing.Point(130, 1180);
            this.button8.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.button8.Name = "button8";
            this.button8.Size = new System.Drawing.Size(320, 37);
            this.button8.TabIndex = 5;
            this.button8.Text = "Get_excel_Sheet_names&Indexes";
            this.button8.UseVisualStyleBackColor = true;
            this.button8.Click += new System.EventHandler(this.button8_Click_1);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.ZnS_M18);
            this.groupBox1.Controls.Add(this.ZnS_M21);
            this.groupBox1.Controls.Add(this.ZnS_M20);
            this.groupBox1.Controls.Add(this.ZnS_M27);
            this.groupBox1.Controls.Add(this.ZnS_M19);
            this.groupBox1.Location = new System.Drawing.Point(27, 755);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(150, 284);
            this.groupBox1.TabIndex = 12;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Första öppning";
            // 
            // ZnS_M18
            // 
            this.ZnS_M18.AutoSize = true;
            this.ZnS_M18.Location = new System.Drawing.Point(6, 39);
            this.ZnS_M18.Name = "ZnS_M18";
            this.ZnS_M18.Size = new System.Drawing.Size(111, 29);
            this.ZnS_M18.TabIndex = 6;
            this.ZnS_M18.Text = "ZnS_M18";
            this.ZnS_M18.UseVisualStyleBackColor = true;
            // 
            // ZnS_M21
            // 
            this.ZnS_M21.AutoSize = true;
            this.ZnS_M21.Location = new System.Drawing.Point(6, 146);
            this.ZnS_M21.Name = "ZnS_M21";
            this.ZnS_M21.Size = new System.Drawing.Size(111, 29);
            this.ZnS_M21.TabIndex = 11;
            this.ZnS_M21.Text = "ZnS_M21";
            this.ZnS_M21.UseVisualStyleBackColor = true;
            // 
            // ZnS_M20
            // 
            this.ZnS_M20.AutoSize = true;
            this.ZnS_M20.Location = new System.Drawing.Point(6, 109);
            this.ZnS_M20.Name = "ZnS_M20";
            this.ZnS_M20.Size = new System.Drawing.Size(111, 29);
            this.ZnS_M20.TabIndex = 8;
            this.ZnS_M20.Text = "ZnS_M20";
            this.ZnS_M20.UseVisualStyleBackColor = true;
            // 
            // ZnS_M27
            // 
            this.ZnS_M27.AutoSize = true;
            this.ZnS_M27.Location = new System.Drawing.Point(6, 181);
            this.ZnS_M27.Name = "ZnS_M27";
            this.ZnS_M27.Size = new System.Drawing.Size(111, 29);
            this.ZnS_M27.TabIndex = 10;
            this.ZnS_M27.Text = "ZnS_M27";
            this.ZnS_M27.UseVisualStyleBackColor = true;
            // 
            // ZnS_M19
            // 
            this.ZnS_M19.AutoSize = true;
            this.ZnS_M19.Location = new System.Drawing.Point(6, 74);
            this.ZnS_M19.Name = "ZnS_M19";
            this.ZnS_M19.Size = new System.Drawing.Size(111, 29);
            this.ZnS_M19.TabIndex = 7;
            this.ZnS_M19.Text = "ZnS_M19";
            this.ZnS_M19.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.ZnS_II_M18);
            this.groupBox2.Controls.Add(this.ZnS_II_M21);
            this.groupBox2.Controls.Add(this.ZnS_II_M20);
            this.groupBox2.Controls.Add(this.ZnS_II_M27);
            this.groupBox2.Controls.Add(this.ZnS_II_M19);
            this.groupBox2.Location = new System.Drawing.Point(197, 755);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(150, 284);
            this.groupBox2.TabIndex = 13;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Andra öppning";
            // 
            // ZnS_II_M18
            // 
            this.ZnS_II_M18.AutoSize = true;
            this.ZnS_II_M18.Location = new System.Drawing.Point(6, 39);
            this.ZnS_II_M18.Name = "ZnS_II_M18";
            this.ZnS_II_M18.Size = new System.Drawing.Size(128, 29);
            this.ZnS_II_M18.TabIndex = 6;
            this.ZnS_II_M18.Text = "ZnS_II_M18";
            this.ZnS_II_M18.UseVisualStyleBackColor = true;
            // 
            // ZnS_II_M21
            // 
            this.ZnS_II_M21.AutoSize = true;
            this.ZnS_II_M21.Location = new System.Drawing.Point(6, 146);
            this.ZnS_II_M21.Name = "ZnS_II_M21";
            this.ZnS_II_M21.Size = new System.Drawing.Size(128, 29);
            this.ZnS_II_M21.TabIndex = 11;
            this.ZnS_II_M21.Text = "ZnS_II_M21";
            this.ZnS_II_M21.UseVisualStyleBackColor = true;
            // 
            // ZnS_II_M20
            // 
            this.ZnS_II_M20.AutoSize = true;
            this.ZnS_II_M20.Location = new System.Drawing.Point(6, 109);
            this.ZnS_II_M20.Name = "ZnS_II_M20";
            this.ZnS_II_M20.Size = new System.Drawing.Size(128, 29);
            this.ZnS_II_M20.TabIndex = 8;
            this.ZnS_II_M20.Text = "ZnS_II_M20";
            this.ZnS_II_M20.UseVisualStyleBackColor = true;
            // 
            // ZnS_II_M27
            // 
            this.ZnS_II_M27.AutoSize = true;
            this.ZnS_II_M27.Location = new System.Drawing.Point(6, 181);
            this.ZnS_II_M27.Name = "ZnS_II_M27";
            this.ZnS_II_M27.Size = new System.Drawing.Size(128, 29);
            this.ZnS_II_M27.TabIndex = 10;
            this.ZnS_II_M27.Text = "ZnS_II_M27";
            this.ZnS_II_M27.UseVisualStyleBackColor = true;
            // 
            // ZnS_II_M19
            // 
            this.ZnS_II_M19.AutoSize = true;
            this.ZnS_II_M19.Location = new System.Drawing.Point(6, 74);
            this.ZnS_II_M19.Name = "ZnS_II_M19";
            this.ZnS_II_M19.Size = new System.Drawing.Size(128, 29);
            this.ZnS_II_M19.TabIndex = 7;
            this.ZnS_II_M19.Text = "ZnS_II_M19";
            this.ZnS_II_M19.UseVisualStyleBackColor = true;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.ZnS_III_M18);
            this.groupBox3.Controls.Add(this.ZnS_III_M21);
            this.groupBox3.Controls.Add(this.ZnS_III_M20);
            this.groupBox3.Controls.Add(this.ZnS_III_M27);
            this.groupBox3.Controls.Add(this.ZnS_III_M19);
            this.groupBox3.Location = new System.Drawing.Point(376, 755);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(150, 284);
            this.groupBox3.TabIndex = 14;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Tredje öppning";
            // 
            // ZnS_III_M18
            // 
            this.ZnS_III_M18.AutoSize = true;
            this.ZnS_III_M18.Location = new System.Drawing.Point(6, 39);
            this.ZnS_III_M18.Name = "ZnS_III_M18";
            this.ZnS_III_M18.Size = new System.Drawing.Size(133, 29);
            this.ZnS_III_M18.TabIndex = 6;
            this.ZnS_III_M18.Text = "ZnS_III_M18";
            this.ZnS_III_M18.UseVisualStyleBackColor = true;
            // 
            // ZnS_III_M21
            // 
            this.ZnS_III_M21.AutoSize = true;
            this.ZnS_III_M21.Location = new System.Drawing.Point(6, 146);
            this.ZnS_III_M21.Name = "ZnS_III_M21";
            this.ZnS_III_M21.Size = new System.Drawing.Size(133, 29);
            this.ZnS_III_M21.TabIndex = 11;
            this.ZnS_III_M21.Text = "ZnS_III_M21";
            this.ZnS_III_M21.UseVisualStyleBackColor = true;
            // 
            // ZnS_III_M20
            // 
            this.ZnS_III_M20.AutoSize = true;
            this.ZnS_III_M20.Location = new System.Drawing.Point(6, 109);
            this.ZnS_III_M20.Name = "ZnS_III_M20";
            this.ZnS_III_M20.Size = new System.Drawing.Size(133, 29);
            this.ZnS_III_M20.TabIndex = 8;
            this.ZnS_III_M20.Text = "ZnS_III_M20";
            this.ZnS_III_M20.UseVisualStyleBackColor = true;
            // 
            // ZnS_III_M27
            // 
            this.ZnS_III_M27.AutoSize = true;
            this.ZnS_III_M27.Location = new System.Drawing.Point(6, 181);
            this.ZnS_III_M27.Name = "ZnS_III_M27";
            this.ZnS_III_M27.Size = new System.Drawing.Size(133, 29);
            this.ZnS_III_M27.TabIndex = 10;
            this.ZnS_III_M27.Text = "ZnS_III_M27";
            this.ZnS_III_M27.UseVisualStyleBackColor = true;
            // 
            // ZnS_III_M19
            // 
            this.ZnS_III_M19.AutoSize = true;
            this.ZnS_III_M19.Location = new System.Drawing.Point(6, 74);
            this.ZnS_III_M19.Name = "ZnS_III_M19";
            this.ZnS_III_M19.Size = new System.Drawing.Size(133, 29);
            this.ZnS_III_M19.TabIndex = 7;
            this.ZnS_III_M19.Text = "ZnS_III_M19";
            this.ZnS_III_M19.UseVisualStyleBackColor = true;
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(33, 12);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersWidth = 62;
            this.dataGridView1.RowTemplate.Height = 33;
            this.dataGridView1.Size = new System.Drawing.Size(482, 725);
            this.dataGridView1.TabIndex = 15;
            // 
            // Export_DT_To_Excel
            // 
            this.Export_DT_To_Excel.Location = new System.Drawing.Point(270, 1236);
            this.Export_DT_To_Excel.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.Export_DT_To_Excel.Name = "Export_DT_To_Excel";
            this.Export_DT_To_Excel.Size = new System.Drawing.Size(180, 37);
            this.Export_DT_To_Excel.TabIndex = 16;
            this.Export_DT_To_Excel.Text = "Export_DT_To_Excel";
            this.Export_DT_To_Excel.UseVisualStyleBackColor = true;
            this.Export_DT_To_Excel.Click += new System.EventHandler(this.Export_DT_To_Excel_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(130, 1236);
            this.button2.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(132, 37);
            this.button2.TabIndex = 17;
            this.button2.Text = "KlistraInD";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click_1);
            // 
            // Form1
            // 
            this.AccessibleName = "";
            this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(551, 1299);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.Export_DT_To_Excel);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.button8);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.Name = "Form1";
            this.Tag = "";
            this.Text = "KS_Automation:TEST";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button8;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.CheckBox ZnS_M18;
        private System.Windows.Forms.CheckBox ZnS_M21;
        private System.Windows.Forms.CheckBox ZnS_M20;
        private System.Windows.Forms.CheckBox ZnS_M27;
        private System.Windows.Forms.CheckBox ZnS_M19;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.CheckBox ZnS_II_M18;
        private System.Windows.Forms.CheckBox ZnS_II_M21;
        private System.Windows.Forms.CheckBox ZnS_II_M20;
        private System.Windows.Forms.CheckBox ZnS_II_M27;
        private System.Windows.Forms.CheckBox ZnS_II_M19;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.CheckBox ZnS_III_M18;
        private System.Windows.Forms.CheckBox ZnS_III_M21;
        private System.Windows.Forms.CheckBox ZnS_III_M20;
        private System.Windows.Forms.CheckBox ZnS_III_M27;
        private System.Windows.Forms.CheckBox ZnS_III_M19;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button Export_DT_To_Excel;
        private System.Windows.Forms.Button button2;
    }
}

