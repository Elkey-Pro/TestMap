namespace TestMap
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
            this.txt_DocxOutPutPath = new System.Windows.Forms.TextBox();
            this.txt_xlsxPath = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label7 = new System.Windows.Forms.Label();
            this.nud_zoom = new System.Windows.Forms.NumericUpDown();
            this.label6 = new System.Windows.Forms.Label();
            this.nud_maxzoom = new System.Windows.Forms.NumericUpDown();
            this.label5 = new System.Windows.Forms.Label();
            this.nud_minzoom = new System.Windows.Forms.NumericUpDown();
            this.button1 = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.txtLong = new System.Windows.Forms.TextBox();
            this.txtLat = new System.Windows.Forms.TextBox();
            this.txtCityName = new System.Windows.Forms.TextBox();
            this.txtCountryName = new System.Windows.Forms.TextBox();
            this.gmap = new GMap.NET.WindowsForms.GMapControl();
            this.button2 = new System.Windows.Forms.Button();
            this.btn_MakeReport = new System.Windows.Forms.Button();
            this.txt_log = new System.Windows.Forms.TextBox();
            this.btn_openOutPut = new System.Windows.Forms.Button();
            this.label10 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nud_zoom)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nud_maxzoom)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nud_minzoom)).BeginInit();
            this.SuspendLayout();
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(61, 4);
            // 
            // txt_DocxOutPutPath
            // 
            this.txt_DocxOutPutPath.Location = new System.Drawing.Point(105, 49);
            this.txt_DocxOutPutPath.Name = "txt_DocxOutPutPath";
            this.txt_DocxOutPutPath.Size = new System.Drawing.Size(166, 20);
            this.txt_DocxOutPutPath.TabIndex = 18;
            // 
            // txt_xlsxPath
            // 
            this.txt_xlsxPath.Location = new System.Drawing.Point(78, 19);
            this.txt_xlsxPath.Name = "txt_xlsxPath";
            this.txt_xlsxPath.Size = new System.Drawing.Size(193, 20);
            this.txt_xlsxPath.TabIndex = 19;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(6, 23);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(51, 13);
            this.label8.TabIndex = 20;
            this.label8.Text = "Xlsx Path";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(6, 52);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(93, 13);
            this.label9.TabIndex = 21;
            this.label9.Text = "Docx OutPut Path";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.button2);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.nud_zoom);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.nud_maxzoom);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.nud_minzoom);
            this.groupBox1.Controls.Add(this.button1);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.txtLong);
            this.groupBox1.Controls.Add(this.txtLat);
            this.groupBox1.Controls.Add(this.txtCityName);
            this.groupBox1.Controls.Add(this.txtCountryName);
            this.groupBox1.Controls.Add(this.gmap);
            this.groupBox1.Location = new System.Drawing.Point(277, 19);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(427, 427);
            this.groupBox1.TabIndex = 22;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Map";
            this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(344, 109);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(34, 13);
            this.label7.TabIndex = 32;
            this.label7.Text = "Zoom";
            this.label7.Click += new System.EventHandler(this.label7_Click);
            // 
            // nud_zoom
            // 
            this.nud_zoom.Location = new System.Drawing.Point(347, 125);
            this.nud_zoom.Name = "nud_zoom";
            this.nud_zoom.Size = new System.Drawing.Size(69, 20);
            this.nud_zoom.TabIndex = 31;
            this.nud_zoom.Value = new decimal(new int[] {
            7,
            0,
            0,
            0});
            this.nud_zoom.ValueChanged += new System.EventHandler(this.nud_zoom_ValueChanged);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(250, 109);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(57, 13);
            this.label6.TabIndex = 30;
            this.label6.Text = "Max Zoom";
            this.label6.Click += new System.EventHandler(this.label6_Click);
            // 
            // nud_maxzoom
            // 
            this.nud_maxzoom.Location = new System.Drawing.Point(253, 125);
            this.nud_maxzoom.Name = "nud_maxzoom";
            this.nud_maxzoom.Size = new System.Drawing.Size(69, 20);
            this.nud_maxzoom.TabIndex = 29;
            this.nud_maxzoom.Value = new decimal(new int[] {
            18,
            0,
            0,
            0});
            this.nud_maxzoom.ValueChanged += new System.EventHandler(this.nud_maxzoom_ValueChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(160, 109);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(54, 13);
            this.label5.TabIndex = 28;
            this.label5.Text = "Min Zoom";
            this.label5.Click += new System.EventHandler(this.label5_Click);
            // 
            // nud_minzoom
            // 
            this.nud_minzoom.Location = new System.Drawing.Point(163, 125);
            this.nud_minzoom.Name = "nud_minzoom";
            this.nud_minzoom.Size = new System.Drawing.Size(69, 20);
            this.nud_minzoom.TabIndex = 27;
            this.nud_minzoom.Value = new decimal(new int[] {
            2,
            0,
            0,
            0});
            this.nud_minzoom.ValueChanged += new System.EventHandler(this.nud_minzoom_ValueChanged);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(29, 99);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(105, 46);
            this.button1.TabIndex = 26;
            this.button1.Text = "Search";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(236, 77);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(31, 13);
            this.label3.TabIndex = 25;
            this.label3.Text = "Long";
            this.label3.Click += new System.EventHandler(this.label3_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(237, 51);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(22, 13);
            this.label4.TabIndex = 24;
            this.label4.Text = "Lat";
            this.label4.Click += new System.EventHandler(this.label4_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(31, 77);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(55, 13);
            this.label2.TabIndex = 23;
            this.label2.Text = "City Name";
            this.label2.Click += new System.EventHandler(this.label2_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(33, 51);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(43, 13);
            this.label1.TabIndex = 22;
            this.label1.Text = "Country";
            // 
            // txtLong
            // 
            this.txtLong.Location = new System.Drawing.Point(297, 74);
            this.txtLong.Name = "txtLong";
            this.txtLong.Size = new System.Drawing.Size(119, 20);
            this.txtLong.TabIndex = 21;
            this.txtLong.TextChanged += new System.EventHandler(this.txtLong_TextChanged);
            // 
            // txtLat
            // 
            this.txtLat.Location = new System.Drawing.Point(297, 48);
            this.txtLat.Name = "txtLat";
            this.txtLat.Size = new System.Drawing.Size(119, 20);
            this.txtLat.TabIndex = 20;
            this.txtLat.TextChanged += new System.EventHandler(this.txtLat_TextChanged);
            // 
            // txtCityName
            // 
            this.txtCityName.Location = new System.Drawing.Point(95, 74);
            this.txtCityName.Name = "txtCityName";
            this.txtCityName.Size = new System.Drawing.Size(119, 20);
            this.txtCityName.TabIndex = 19;
            this.txtCityName.TextChanged += new System.EventHandler(this.txtCityName_TextChanged);
            // 
            // txtCountryName
            // 
            this.txtCountryName.Location = new System.Drawing.Point(95, 48);
            this.txtCountryName.Name = "txtCountryName";
            this.txtCountryName.Size = new System.Drawing.Size(119, 20);
            this.txtCountryName.TabIndex = 18;
            this.txtCountryName.TextChanged += new System.EventHandler(this.txtCountryName_TextChanged);
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
            this.gmap.Location = new System.Drawing.Point(33, 151);
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
            this.gmap.Size = new System.Drawing.Size(383, 231);
            this.gmap.TabIndex = 17;
            this.gmap.Zoom = 13D;
            this.gmap.Load += new System.EventHandler(this.gmap_Load);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(312, 385);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(104, 42);
            this.button2.TabIndex = 33;
            this.button2.Text = "Export Image";
            this.button2.UseVisualStyleBackColor = true;
            // 
            // btn_MakeReport
            // 
            this.btn_MakeReport.Location = new System.Drawing.Point(151, 84);
            this.btn_MakeReport.Name = "btn_MakeReport";
            this.btn_MakeReport.Size = new System.Drawing.Size(120, 37);
            this.btn_MakeReport.TabIndex = 23;
            this.btn_MakeReport.Text = "Run Report";
            this.btn_MakeReport.UseVisualStyleBackColor = true;
            // 
            // txt_log
            // 
            this.txt_log.BackColor = System.Drawing.SystemColors.ControlLight;
            this.txt_log.Location = new System.Drawing.Point(9, 170);
            this.txt_log.Multiline = true;
            this.txt_log.Name = "txt_log";
            this.txt_log.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txt_log.Size = new System.Drawing.Size(262, 231);
            this.txt_log.TabIndex = 24;
            this.txt_log.TextChanged += new System.EventHandler(this.textBox3_TextChanged);
            // 
            // btn_openOutPut
            // 
            this.btn_openOutPut.Location = new System.Drawing.Point(151, 404);
            this.btn_openOutPut.Name = "btn_openOutPut";
            this.btn_openOutPut.Size = new System.Drawing.Size(120, 37);
            this.btn_openOutPut.TabIndex = 25;
            this.btn_openOutPut.Text = "Open OutPut Folder";
            this.btn_openOutPut.UseVisualStyleBackColor = true;
            this.btn_openOutPut.Click += new System.EventHandler(this.button4_Click);
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(12, 154);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(25, 13);
            this.label10.TabIndex = 26;
            this.label10.Text = "Log";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(714, 451);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.btn_openOutPut);
            this.Controls.Add(this.txt_log);
            this.Controls.Add(this.btn_MakeReport);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.txt_xlsxPath);
            this.Controls.Add(this.txt_DocxOutPutPath);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Name = "MainForm";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nud_zoom)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nud_maxzoom)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nud_minzoom)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.TextBox txt_DocxOutPutPath;
        private System.Windows.Forms.TextBox txt_xlsxPath;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.NumericUpDown nud_zoom;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.NumericUpDown nud_maxzoom;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.NumericUpDown nud_minzoom;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtLong;
        private System.Windows.Forms.TextBox txtLat;
        private System.Windows.Forms.TextBox txtCityName;
        private System.Windows.Forms.TextBox txtCountryName;
        private GMap.NET.WindowsForms.GMapControl gmap;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button btn_MakeReport;
        private System.Windows.Forms.TextBox txt_log;
        private System.Windows.Forms.Button btn_openOutPut;
        private System.Windows.Forms.Label label10;
    }
}

