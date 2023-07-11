namespace RvAutoReport
{
    partial class InputLatLong
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
            this.lv_latlong = new System.Windows.Forms.ListView();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.clm_Lat = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.clm_long = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.lbCount = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // lv_latlong
            // 
            this.lv_latlong.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.clm_Lat,
            this.clm_long});
            this.lv_latlong.FullRowSelect = true;
            this.lv_latlong.GridLines = true;
            this.lv_latlong.HideSelection = false;
            this.lv_latlong.Location = new System.Drawing.Point(12, 30);
            this.lv_latlong.Name = "lv_latlong";
            this.lv_latlong.Size = new System.Drawing.Size(213, 275);
            this.lv_latlong.TabIndex = 0;
            this.lv_latlong.UseCompatibleStateImageBehavior = false;
            this.lv_latlong.View = System.Windows.Forms.View.Details;
            this.lv_latlong.KeyUp += new System.Windows.Forms.KeyEventHandler(this.lv_latlong_KeyUp);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(12, 311);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(139, 30);
            this.button1.TabIndex = 1;
            this.button1.Text = "OK";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(157, 311);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(68, 30);
            this.button2.TabIndex = 2;
            this.button2.Text = "Cancel";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // clm_Lat
            // 
            this.clm_Lat.Text = "LAT";
            this.clm_Lat.Width = 105;
            // 
            // clm_long
            // 
            this.clm_long.Text = "LONG";
            this.clm_long.Width = 105;
            // 
            // lbCount
            // 
            this.lbCount.AutoSize = true;
            this.lbCount.Location = new System.Drawing.Point(173, 14);
            this.lbCount.Name = "lbCount";
            this.lbCount.Size = new System.Drawing.Size(40, 13);
            this.lbCount.TabIndex = 3;
            this.lbCount.Text = "Total : ";
            // 
            // InputLatLong
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(235, 346);
            this.ControlBox = false;
            this.Controls.Add(this.lbCount);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.lv_latlong);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Name = "InputLatLong";
            this.Text = "InputLatLong";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListView lv_latlong;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.ColumnHeader clm_Lat;
        private System.Windows.Forms.ColumnHeader clm_long;
        private System.Windows.Forms.Label lbCount;
    }
}