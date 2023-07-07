namespace RvAutoReport
{
    partial class MainForm
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
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.txt_UserName = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txt_docxTemp = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.gmap = new GMap.NET.WindowsForms.GMapControl();
            this.button4 = new System.Windows.Forms.Button();
            this.label10 = new System.Windows.Forms.Label();
            this.btn_openOutPut = new System.Windows.Forms.Button();
            this.txt_log = new System.Windows.Forms.TextBox();
            this.btn_MakeReport = new System.Windows.Forms.Button();
            this.label9 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.txt_xlsxPath = new System.Windows.Forms.TextBox();
            this.txt_DocxOutPutPath = new System.Windows.Forms.TextBox();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.button3 = new System.Windows.Forms.Button();
            this.dataGridView2 = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dGVDataWFromXl = new System.Windows.Forms.DataGridView();
            this.btn_clearLog = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dGVDataWFromXl)).BeginInit();
            this.SuspendLayout();
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(61, 4);
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(801, 513);
            this.tabControl1.TabIndex = 27;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.button1);
            this.tabPage1.Controls.Add(this.btn_clearLog);
            this.tabPage1.Controls.Add(this.txt_UserName);
            this.tabPage1.Controls.Add(this.label2);
            this.tabPage1.Controls.Add(this.txt_docxTemp);
            this.tabPage1.Controls.Add(this.label1);
            this.tabPage1.Controls.Add(this.groupBox1);
            this.tabPage1.Controls.Add(this.button4);
            this.tabPage1.Controls.Add(this.label10);
            this.tabPage1.Controls.Add(this.btn_openOutPut);
            this.tabPage1.Controls.Add(this.txt_log);
            this.tabPage1.Controls.Add(this.btn_MakeReport);
            this.tabPage1.Controls.Add(this.label9);
            this.tabPage1.Controls.Add(this.label8);
            this.tabPage1.Controls.Add(this.txt_xlsxPath);
            this.tabPage1.Controls.Add(this.txt_DocxOutPutPath);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(793, 487);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Main";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // txt_UserName
            // 
            this.txt_UserName.Location = new System.Drawing.Point(107, 87);
            this.txt_UserName.Name = "txt_UserName";
            this.txt_UserName.Size = new System.Drawing.Size(111, 20);
            this.txt_UserName.TabIndex = 41;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(9, 90);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(57, 13);
            this.label2.TabIndex = 40;
            this.label2.Text = "UserName";
            // 
            // txt_docxTemp
            // 
            this.txt_docxTemp.Location = new System.Drawing.Point(107, 36);
            this.txt_docxTemp.Name = "txt_docxTemp";
            this.txt_docxTemp.Size = new System.Drawing.Size(112, 20);
            this.txt_docxTemp.TabIndex = 39;
            this.txt_docxTemp.Text = "RV Rental Tampa Report-v0.3.docx";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 39);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(84, 13);
            this.label1.TabIndex = 38;
            this.label1.Text = "Doc temp Name";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.gmap);
            this.groupBox1.Location = new System.Drawing.Point(225, 10);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(570, 477);
            this.groupBox1.TabIndex = 37;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Map";
            this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
            // 
            // gmap
            // 
            this.gmap.Bearing = 0F;
            this.gmap.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.gmap.CanDragMap = true;
            this.gmap.EmptyTileColor = System.Drawing.Color.Navy;
            this.gmap.GrayScaleMode = false;
            this.gmap.HelperLineOption = GMap.NET.WindowsForms.HelperLineOptions.DontShow;
            this.gmap.LevelsKeepInMemory = 5;
            this.gmap.Location = new System.Drawing.Point(6, 19);
            this.gmap.MarkersEnabled = true;
            this.gmap.MaxZoom = 18;
            this.gmap.MinZoom = 2;
            this.gmap.MouseWheelZoomEnabled = true;
            this.gmap.MouseWheelZoomType = GMap.NET.MouseWheelZoomType.MousePositionAndCenter;
            this.gmap.Name = "gmap";
            this.gmap.NegativeMode = false;
            this.gmap.PolygonsEnabled = true;
            this.gmap.RetryLoadTile = 0;
            this.gmap.RoutesEnabled = true;
            this.gmap.ScaleMode = GMap.NET.WindowsForms.ScaleModes.Integer;
            this.gmap.SelectedAreaFillColor = System.Drawing.Color.FromArgb(((int)(((byte)(33)))), ((int)(((byte)(65)))), ((int)(((byte)(105)))), ((int)(((byte)(225)))));
            this.gmap.ShowTileGridLines = false;
            this.gmap.Size = new System.Drawing.Size(552, 452);
            this.gmap.TabIndex = 17;
            this.gmap.Zoom = 13D;
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(107, 111);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(112, 44);
            this.button4.TabIndex = 28;
            this.button4.Text = "Run Report2";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(9, 188);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(25, 13);
            this.label10.TabIndex = 34;
            this.label10.Text = "Log";
            this.label10.Click += new System.EventHandler(this.label10_Click);
            // 
            // btn_openOutPut
            // 
            this.btn_openOutPut.Location = new System.Drawing.Point(99, 441);
            this.btn_openOutPut.Name = "btn_openOutPut";
            this.btn_openOutPut.Size = new System.Drawing.Size(120, 37);
            this.btn_openOutPut.TabIndex = 33;
            this.btn_openOutPut.Text = "Open OutPut Folder";
            this.btn_openOutPut.UseVisualStyleBackColor = true;
            this.btn_openOutPut.Click += new System.EventHandler(this.btn_openOutPut_Click);
            // 
            // txt_log
            // 
            this.txt_log.BackColor = System.Drawing.SystemColors.ControlLight;
            this.txt_log.Location = new System.Drawing.Point(12, 204);
            this.txt_log.Multiline = true;
            this.txt_log.Name = "txt_log";
            this.txt_log.ReadOnly = true;
            this.txt_log.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txt_log.Size = new System.Drawing.Size(207, 231);
            this.txt_log.TabIndex = 32;
            // 
            // btn_MakeReport
            // 
            this.btn_MakeReport.Location = new System.Drawing.Point(107, 161);
            this.btn_MakeReport.Name = "btn_MakeReport";
            this.btn_MakeReport.Size = new System.Drawing.Size(112, 37);
            this.btn_MakeReport.TabIndex = 31;
            this.btn_MakeReport.Text = "Run Report";
            this.btn_MakeReport.UseVisualStyleBackColor = true;
            this.btn_MakeReport.Visible = false;
            this.btn_MakeReport.Click += new System.EventHandler(this.btn_MakeReport_Click_1);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(8, 66);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(93, 13);
            this.label9.TabIndex = 30;
            this.label9.Text = "Docx OutPut Path";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(9, 10);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(51, 13);
            this.label8.TabIndex = 29;
            this.label8.Text = "Xlsx Path";
            // 
            // txt_xlsxPath
            // 
            this.txt_xlsxPath.Location = new System.Drawing.Point(81, 6);
            this.txt_xlsxPath.Name = "txt_xlsxPath";
            this.txt_xlsxPath.Size = new System.Drawing.Size(138, 20);
            this.txt_xlsxPath.TabIndex = 28;
            // 
            // txt_DocxOutPutPath
            // 
            this.txt_DocxOutPutPath.Location = new System.Drawing.Point(107, 63);
            this.txt_DocxOutPutPath.Name = "txt_DocxOutPutPath";
            this.txt_DocxOutPutPath.Size = new System.Drawing.Size(111, 20);
            this.txt_DocxOutPutPath.TabIndex = 27;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.button3);
            this.tabPage2.Controls.Add(this.dataGridView2);
            this.tabPage2.Controls.Add(this.dGVDataWFromXl);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(793, 487);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Setting";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(269, 6);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(88, 30);
            this.button3.TabIndex = 2;
            this.button3.Text = "SAVE";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // dataGridView2
            // 
            this.dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView2.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn1,
            this.dataGridViewTextBoxColumn2});
            this.dataGridView2.Location = new System.Drawing.Point(423, 42);
            this.dataGridView2.Name = "dataGridView2";
            this.dataGridView2.Size = new System.Drawing.Size(361, 441);
            this.dataGridView2.TabIndex = 1;
            // 
            // dataGridViewTextBoxColumn1
            // 
            this.dataGridViewTextBoxColumn1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.dataGridViewTextBoxColumn1.HeaderText = "String To Find";
            this.dataGridViewTextBoxColumn1.Name = "dataGridViewTextBoxColumn1";
            this.dataGridViewTextBoxColumn1.Width = 98;
            // 
            // dataGridViewTextBoxColumn2
            // 
            this.dataGridViewTextBoxColumn2.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader;
            this.dataGridViewTextBoxColumn2.HeaderText = "Cell Location";
            this.dataGridViewTextBoxColumn2.Name = "dataGridViewTextBoxColumn2";
            this.dataGridViewTextBoxColumn2.Width = 93;
            // 
            // dGVDataWFromXl
            // 
            this.dGVDataWFromXl.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dGVDataWFromXl.Location = new System.Drawing.Point(0, 42);
            this.dGVDataWFromXl.Name = "dGVDataWFromXl";
            this.dGVDataWFromXl.Size = new System.Drawing.Size(357, 438);
            this.dGVDataWFromXl.TabIndex = 0;
            // 
            // btn_clearLog
            // 
            this.btn_clearLog.Location = new System.Drawing.Point(11, 441);
            this.btn_clearLog.Name = "btn_clearLog";
            this.btn_clearLog.Size = new System.Drawing.Size(75, 23);
            this.btn_clearLog.TabIndex = 18;
            this.btn_clearLog.Text = "ClearLog";
            this.btn_clearLog.UseVisualStyleBackColor = true;
            this.btn_clearLog.Click += new System.EventHandler(this.btn_clearLog_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(8, 111);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 44);
            this.button1.TabIndex = 42;
            this.button1.Text = "Run Report1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(803, 510);
            this.Controls.Add(this.tabControl1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Name = "MainForm";
            this.Text = "Auto Report";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dGVDataWFromXl)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Button btn_openOutPut;
        private System.Windows.Forms.TextBox txt_log;
        private System.Windows.Forms.Button btn_MakeReport;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox txt_xlsxPath;
        private System.Windows.Forms.TextBox txt_DocxOutPutPath;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.DataGridView dataGridView2;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
        private System.Windows.Forms.DataGridView dGVDataWFromXl;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.GroupBox groupBox1;
        private GMap.NET.WindowsForms.GMapControl gmap;
        private System.Windows.Forms.TextBox txt_UserName;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txt_docxTemp;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btn_clearLog;
        private System.Windows.Forms.Button button1;
    }
}

