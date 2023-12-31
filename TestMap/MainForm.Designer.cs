﻿namespace RvAutoReport
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.button6 = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.btn_clearLog = new System.Windows.Forms.Button();
            this.txt_UserName = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.gmap = new GMap.NET.WindowsForms.GMapControl();
            this.button4 = new System.Windows.Forms.Button();
            this.label10 = new System.Windows.Forms.Label();
            this.btn_openOutPut = new System.Windows.Forms.Button();
            this.txt_log = new System.Windows.Forms.TextBox();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btn_OpenTemp = new System.Windows.Forms.Button();
            this.btn_openLogo = new System.Windows.Forms.Button();
            this.btn_openDocx = new System.Windows.Forms.Button();
            this.btn_openXlsx = new System.Windows.Forms.Button();
            this.btn_openCSV = new System.Windows.Forms.Button();
            this.label12 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.btn_LoadLogoPath = new System.Windows.Forms.Button();
            this.btn_LoadDocxPath = new System.Windows.Forms.Button();
            this.btn_LoadXlsxPath = new System.Windows.Forms.Button();
            this.btn_LoadCsvPath = new System.Windows.Forms.Button();
            this.label6 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.txt_circleDiameter = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txt_logopath = new System.Windows.Forms.TextBox();
            this.txt_docxOutPut = new System.Windows.Forms.TextBox();
            this.txt_xlsxOutput = new System.Windows.Forms.TextBox();
            this.txt_csvInput = new System.Windows.Forms.TextBox();
            this.cbb_SelectReport = new System.Windows.Forms.ComboBox();
            this.rd_runOneReport = new System.Windows.Forms.RadioButton();
            this.rd_runAllReport = new System.Windows.Forms.RadioButton();
            this.button3 = new System.Windows.Forms.Button();
            this.dGVDataWFromXl = new System.Windows.Forms.DataGridView();
            this.btn_fullauto = new System.Windows.Forms.Button();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.groupBox2.SuspendLayout();
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
            this.tabPage1.Controls.Add(this.btn_fullauto);
            this.tabPage1.Controls.Add(this.button6);
            this.tabPage1.Controls.Add(this.button5);
            this.tabPage1.Controls.Add(this.button2);
            this.tabPage1.Controls.Add(this.button1);
            this.tabPage1.Controls.Add(this.btn_clearLog);
            this.tabPage1.Controls.Add(this.txt_UserName);
            this.tabPage1.Controls.Add(this.label2);
            this.tabPage1.Controls.Add(this.groupBox1);
            this.tabPage1.Controls.Add(this.button4);
            this.tabPage1.Controls.Add(this.label10);
            this.tabPage1.Controls.Add(this.btn_openOutPut);
            this.tabPage1.Controls.Add(this.txt_log);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(793, 487);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Main";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // button6
            // 
            this.button6.Location = new System.Drawing.Point(11, 36);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(208, 44);
            this.button6.TabIndex = 48;
            this.button6.Text = "Run Phase 1 : Csv To Xlsx";
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Click += new System.EventHandler(this.button6_Click);
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(142, 209);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(75, 23);
            this.button5.TabIndex = 47;
            this.button5.Text = "Clear Map";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(142, 180);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 43;
            this.button2.Text = "Kill Excel";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(99, 135);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(120, 39);
            this.button1.TabIndex = 42;
            this.button1.Text = "Run Custom LatLong";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
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
            // txt_UserName
            // 
            this.txt_UserName.Location = new System.Drawing.Point(106, 10);
            this.txt_UserName.Name = "txt_UserName";
            this.txt_UserName.Size = new System.Drawing.Size(111, 20);
            this.txt_UserName.TabIndex = 41;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(8, 13);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(95, 13);
            this.label2.TabIndex = 40;
            this.label2.Text = "Report User Name";
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
            this.button4.Location = new System.Drawing.Point(12, 86);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(207, 43);
            this.button4.TabIndex = 28;
            this.button4.Text = "Run Phase 2 : Xlsx To Word Report";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(9, 219);
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
            this.txt_log.Location = new System.Drawing.Point(12, 238);
            this.txt_log.Multiline = true;
            this.txt_log.Name = "txt_log";
            this.txt_log.ReadOnly = true;
            this.txt_log.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txt_log.Size = new System.Drawing.Size(207, 197);
            this.txt_log.TabIndex = 32;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.groupBox2);
            this.tabPage2.Controls.Add(this.dGVDataWFromXl);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(793, 487);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Setting";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.btn_OpenTemp);
            this.groupBox2.Controls.Add(this.btn_openLogo);
            this.groupBox2.Controls.Add(this.btn_openDocx);
            this.groupBox2.Controls.Add(this.btn_openXlsx);
            this.groupBox2.Controls.Add(this.btn_openCSV);
            this.groupBox2.Controls.Add(this.label12);
            this.groupBox2.Controls.Add(this.label11);
            this.groupBox2.Controls.Add(this.label7);
            this.groupBox2.Controls.Add(this.btn_LoadLogoPath);
            this.groupBox2.Controls.Add(this.btn_LoadDocxPath);
            this.groupBox2.Controls.Add(this.btn_LoadXlsxPath);
            this.groupBox2.Controls.Add(this.btn_LoadCsvPath);
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Controls.Add(this.txt_circleDiameter);
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.txt_logopath);
            this.groupBox2.Controls.Add(this.txt_docxOutPut);
            this.groupBox2.Controls.Add(this.txt_xlsxOutput);
            this.groupBox2.Controls.Add(this.txt_csvInput);
            this.groupBox2.Controls.Add(this.cbb_SelectReport);
            this.groupBox2.Controls.Add(this.rd_runOneReport);
            this.groupBox2.Controls.Add(this.rd_runAllReport);
            this.groupBox2.Controls.Add(this.button3);
            this.groupBox2.Location = new System.Drawing.Point(363, 42);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(434, 439);
            this.groupBox2.TabIndex = 3;
            this.groupBox2.TabStop = false;
            // 
            // btn_OpenTemp
            // 
            this.btn_OpenTemp.Location = new System.Drawing.Point(383, 61);
            this.btn_OpenTemp.Name = "btn_OpenTemp";
            this.btn_OpenTemp.Size = new System.Drawing.Size(41, 23);
            this.btn_OpenTemp.TabIndex = 62;
            this.btn_OpenTemp.Text = "Open";
            this.btn_OpenTemp.UseVisualStyleBackColor = true;
            this.btn_OpenTemp.Click += new System.EventHandler(this.btn_OpenTemp_Click);
            // 
            // btn_openLogo
            // 
            this.btn_openLogo.Location = new System.Drawing.Point(383, 224);
            this.btn_openLogo.Name = "btn_openLogo";
            this.btn_openLogo.Size = new System.Drawing.Size(41, 23);
            this.btn_openLogo.TabIndex = 61;
            this.btn_openLogo.Text = "Open";
            this.btn_openLogo.UseVisualStyleBackColor = true;
            this.btn_openLogo.Click += new System.EventHandler(this.btn_openLogo_Click);
            // 
            // btn_openDocx
            // 
            this.btn_openDocx.Location = new System.Drawing.Point(383, 184);
            this.btn_openDocx.Name = "btn_openDocx";
            this.btn_openDocx.Size = new System.Drawing.Size(41, 23);
            this.btn_openDocx.TabIndex = 60;
            this.btn_openDocx.Text = "Open";
            this.btn_openDocx.UseVisualStyleBackColor = true;
            this.btn_openDocx.Click += new System.EventHandler(this.btn_openDocx_Click);
            // 
            // btn_openXlsx
            // 
            this.btn_openXlsx.Location = new System.Drawing.Point(383, 141);
            this.btn_openXlsx.Name = "btn_openXlsx";
            this.btn_openXlsx.Size = new System.Drawing.Size(41, 23);
            this.btn_openXlsx.TabIndex = 59;
            this.btn_openXlsx.Text = "Open";
            this.btn_openXlsx.UseVisualStyleBackColor = true;
            this.btn_openXlsx.Click += new System.EventHandler(this.btn_openXlsx_Click);
            // 
            // btn_openCSV
            // 
            this.btn_openCSV.Location = new System.Drawing.Point(383, 98);
            this.btn_openCSV.Name = "btn_openCSV";
            this.btn_openCSV.Size = new System.Drawing.Size(41, 23);
            this.btn_openCSV.TabIndex = 58;
            this.btn_openCSV.Text = "Open";
            this.btn_openCSV.UseVisualStyleBackColor = true;
            this.btn_openCSV.Click += new System.EventHandler(this.btn_openCSV_Click);
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(6, 188);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(92, 13);
            this.label12.TabIndex = 57;
            this.label12.Text = "Docx Output Path";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(6, 147);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(86, 13);
            this.label11.TabIndex = 56;
            this.label11.Text = "Xlsx Output Path";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(6, 229);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(56, 13);
            this.label7.TabIndex = 55;
            this.label7.Text = "Logo Path";
            // 
            // btn_LoadLogoPath
            // 
            this.btn_LoadLogoPath.Location = new System.Drawing.Point(133, 223);
            this.btn_LoadLogoPath.Name = "btn_LoadLogoPath";
            this.btn_LoadLogoPath.Size = new System.Drawing.Size(31, 23);
            this.btn_LoadLogoPath.TabIndex = 54;
            this.btn_LoadLogoPath.Text = "...";
            this.btn_LoadLogoPath.UseVisualStyleBackColor = true;
            this.btn_LoadLogoPath.Click += new System.EventHandler(this.btn_LoadLogoPath_Click);
            // 
            // btn_LoadDocxPath
            // 
            this.btn_LoadDocxPath.Location = new System.Drawing.Point(133, 183);
            this.btn_LoadDocxPath.Name = "btn_LoadDocxPath";
            this.btn_LoadDocxPath.Size = new System.Drawing.Size(31, 23);
            this.btn_LoadDocxPath.TabIndex = 53;
            this.btn_LoadDocxPath.Text = "...";
            this.btn_LoadDocxPath.UseVisualStyleBackColor = true;
            this.btn_LoadDocxPath.Click += new System.EventHandler(this.btn_LoadDocxPath_Click);
            // 
            // btn_LoadXlsxPath
            // 
            this.btn_LoadXlsxPath.Location = new System.Drawing.Point(133, 141);
            this.btn_LoadXlsxPath.Name = "btn_LoadXlsxPath";
            this.btn_LoadXlsxPath.Size = new System.Drawing.Size(31, 23);
            this.btn_LoadXlsxPath.TabIndex = 52;
            this.btn_LoadXlsxPath.Text = "...";
            this.btn_LoadXlsxPath.UseVisualStyleBackColor = true;
            this.btn_LoadXlsxPath.Click += new System.EventHandler(this.btn_LoadXlsxPath_Click);
            // 
            // btn_LoadCsvPath
            // 
            this.btn_LoadCsvPath.Location = new System.Drawing.Point(133, 98);
            this.btn_LoadCsvPath.Name = "btn_LoadCsvPath";
            this.btn_LoadCsvPath.Size = new System.Drawing.Size(31, 23);
            this.btn_LoadCsvPath.TabIndex = 51;
            this.btn_LoadCsvPath.Text = "...";
            this.btn_LoadCsvPath.UseVisualStyleBackColor = true;
            this.btn_LoadCsvPath.Click += new System.EventHandler(this.btn_LoadCsvPath_Click);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(6, 100);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(80, 13);
            this.label6.TabIndex = 50;
            this.label6.Text = "CSV Input Path";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(230, 272);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(31, 13);
            this.label4.TabIndex = 49;
            this.label4.Text = "Miles";
            this.label4.Click += new System.EventHandler(this.label4_Click);
            // 
            // txt_circleDiameter
            // 
            this.txt_circleDiameter.Location = new System.Drawing.Point(173, 265);
            this.txt_circleDiameter.Name = "txt_circleDiameter";
            this.txt_circleDiameter.Size = new System.Drawing.Size(51, 20);
            this.txt_circleDiameter.TabIndex = 48;
            this.txt_circleDiameter.Text = "50";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(6, 268);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(78, 13);
            this.label3.TabIndex = 47;
            this.label3.Text = "Circle Diameter";
            this.label3.Click += new System.EventHandler(this.label3_Click);
            // 
            // txt_logopath
            // 
            this.txt_logopath.Location = new System.Drawing.Point(173, 226);
            this.txt_logopath.Name = "txt_logopath";
            this.txt_logopath.Size = new System.Drawing.Size(204, 20);
            this.txt_logopath.TabIndex = 10;
            this.txt_logopath.TextChanged += new System.EventHandler(this.txt_logopath_TextChanged);
            // 
            // txt_docxOutPut
            // 
            this.txt_docxOutPut.Location = new System.Drawing.Point(173, 186);
            this.txt_docxOutPut.Name = "txt_docxOutPut";
            this.txt_docxOutPut.Size = new System.Drawing.Size(204, 20);
            this.txt_docxOutPut.TabIndex = 9;
            this.txt_docxOutPut.TextChanged += new System.EventHandler(this.txt_docxOutPut_TextChanged);
            // 
            // txt_xlsxOutput
            // 
            this.txt_xlsxOutput.Location = new System.Drawing.Point(173, 144);
            this.txt_xlsxOutput.Name = "txt_xlsxOutput";
            this.txt_xlsxOutput.Size = new System.Drawing.Size(204, 20);
            this.txt_xlsxOutput.TabIndex = 8;
            this.txt_xlsxOutput.TextChanged += new System.EventHandler(this.txt_xlsxOutput_TextChanged);
            // 
            // txt_csvInput
            // 
            this.txt_csvInput.Location = new System.Drawing.Point(173, 100);
            this.txt_csvInput.Name = "txt_csvInput";
            this.txt_csvInput.Size = new System.Drawing.Size(204, 20);
            this.txt_csvInput.TabIndex = 7;
            this.txt_csvInput.TextChanged += new System.EventHandler(this.txt_csvInput_TextChanged);
            // 
            // cbb_SelectReport
            // 
            this.cbb_SelectReport.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbb_SelectReport.FormattingEnabled = true;
            this.cbb_SelectReport.Location = new System.Drawing.Point(133, 60);
            this.cbb_SelectReport.Name = "cbb_SelectReport";
            this.cbb_SelectReport.Size = new System.Drawing.Size(244, 21);
            this.cbb_SelectReport.TabIndex = 6;
            this.cbb_SelectReport.SelectedIndexChanged += new System.EventHandler(this.cbb_SelectReport_SelectedIndexChanged);
            this.cbb_SelectReport.MouseDown += new System.Windows.Forms.MouseEventHandler(this.cbb_SelectReport_MouseDown);
            // 
            // rd_runOneReport
            // 
            this.rd_runOneReport.AutoSize = true;
            this.rd_runOneReport.Location = new System.Drawing.Point(6, 61);
            this.rd_runOneReport.Name = "rd_runOneReport";
            this.rd_runOneReport.Size = new System.Drawing.Size(121, 17);
            this.rd_runOneReport.TabIndex = 5;
            this.rd_runOneReport.TabStop = true;
            this.rd_runOneReport.Text = "Run Specific Report";
            this.rd_runOneReport.UseVisualStyleBackColor = true;
            this.rd_runOneReport.CheckedChanged += new System.EventHandler(this.rd_runOneReport_CheckedChanged);
            // 
            // rd_runAllReport
            // 
            this.rd_runAllReport.AutoSize = true;
            this.rd_runAllReport.Location = new System.Drawing.Point(6, 29);
            this.rd_runAllReport.Name = "rd_runAllReport";
            this.rd_runAllReport.Size = new System.Drawing.Size(94, 17);
            this.rd_runAllReport.TabIndex = 4;
            this.rd_runAllReport.TabStop = true;
            this.rd_runAllReport.Text = "Run All Report";
            this.rd_runAllReport.UseVisualStyleBackColor = true;
            this.rd_runAllReport.CheckedChanged += new System.EventHandler(this.rd_runAllReport_CheckedChanged);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(313, 395);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(114, 43);
            this.button3.TabIndex = 3;
            this.button3.Text = "SAVE";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // dGVDataWFromXl
            // 
            this.dGVDataWFromXl.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dGVDataWFromXl.Location = new System.Drawing.Point(0, 42);
            this.dGVDataWFromXl.Name = "dGVDataWFromXl";
            this.dGVDataWFromXl.Size = new System.Drawing.Size(357, 438);
            this.dGVDataWFromXl.TabIndex = 0;
            // 
            // btn_fullauto
            // 
            this.btn_fullauto.Location = new System.Drawing.Point(12, 135);
            this.btn_fullauto.Name = "btn_fullauto";
            this.btn_fullauto.Size = new System.Drawing.Size(82, 39);
            this.btn_fullauto.TabIndex = 49;
            this.btn_fullauto.Text = "Full Auto";
            this.btn_fullauto.UseVisualStyleBackColor = true;
            this.btn_fullauto.Click += new System.EventHandler(this.btn_fullauto_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(803, 510);
            this.Controls.Add(this.tabControl1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "MainForm";
            this.Text = "Auto Report";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.MainForm_FormClosed);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
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
        private System.Windows.Forms.DataGridView dGVDataWFromXl;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.GroupBox groupBox1;
        private GMap.NET.WindowsForms.GMapControl gmap;
        private System.Windows.Forms.TextBox txt_UserName;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btn_clearLog;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txt_circleDiameter;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txt_logopath;
        private System.Windows.Forms.TextBox txt_docxOutPut;
        private System.Windows.Forms.TextBox txt_xlsxOutput;
        private System.Windows.Forms.TextBox txt_csvInput;
        private System.Windows.Forms.ComboBox cbb_SelectReport;
        private System.Windows.Forms.RadioButton rd_runOneReport;
        private System.Windows.Forms.RadioButton rd_runAllReport;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Button btn_LoadLogoPath;
        private System.Windows.Forms.Button btn_LoadDocxPath;
        private System.Windows.Forms.Button btn_LoadXlsxPath;
        private System.Windows.Forms.Button btn_LoadCsvPath;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button btn_openLogo;
        private System.Windows.Forms.Button btn_openDocx;
        private System.Windows.Forms.Button btn_openXlsx;
        private System.Windows.Forms.Button btn_openCSV;
        private System.Windows.Forms.Button btn_OpenTemp;
        private System.Windows.Forms.Button btn_fullauto;
    }
}

