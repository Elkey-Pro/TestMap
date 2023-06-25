﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using GMap;
using GMap.NET.WindowsForms;

namespace TestMap
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }
        public string ExcelPath = @"C:\Personal\TDQ\PJ\FREEL\cefSharpTest\Insert Image\Insert Image\bin\x86\Debug\XLSX_OUTPUT\FL-Miami-6-7-23-1023AM.xlsx";
        public DataTable Excel_data;

        private void Form1_Load(object sender, EventArgs e)
        {
            txt_xlsxPath.Text = Environment.CurrentDirectory + "XLSX_OUTPUT" ;
            txt_DocxOutPutPath.Text = Environment.CurrentDirectory + "WORD_OUTPUT";

        }

        private void button1_Click(object sender, EventArgs e)
        {
            gmap.MaxZoom = (int)(nud_maxzoom.Value);
            gmap.MinZoom = (int)(nud_minzoom.Value);
            gmap.Zoom = (int)(nud_zoom.Value);
            gmap.MapProvider = GMap.NET.MapProviders.GoogleMapProvider.Instance;
            GMap.NET.GMaps.Instance.Mode = GMap.NET.AccessMode.ServerOnly;
            gmap.Position = new GMap.NET.PointLatLng(double.Parse(txtLat.Text), double.Parse(txtLong.Text));


            DataTable newTable = new DataTable("NewTable");
            newTable.Columns.Add("CityName", typeof(string));
            newTable.Columns.Add("lat", typeof(double));
            newTable.Columns.Add("long", typeof(double));

            foreach (DataRow row in Excel_data.Rows)
            {
                string City = (string)row["CITY"];
                double lat = (double)row["LATITUDE"];
                double lng = (double)row["LONGITUDE"];
                newTable.Rows.Add(City,lat, lng);
            }

            foreach(DataRow dtrow in newTable.Rows)
            {
                GMap.NET.WindowsForms.GMapOverlay markers = new GMap.NET.WindowsForms.GMapOverlay("markers");
                GMap.NET.WindowsForms.GMapMarker marker = new GMap.NET.WindowsForms.Markers.GMarkerGoogle(
                                                                new GMap.NET.PointLatLng(dtrow.Field<double>("lat"), dtrow.Field<double>("long")),
                                                                            GMap.NET.WindowsForms.Markers.GMarkerGoogleType.red_small);
                markers.Markers.Add(marker);
                gmap.Overlays.Add(markers);
            }

           
        }


        private static DataTable ReadExcelFile(string sheetName, string path)
        {
            using (OleDbConnection conn = new OleDbConnection())
            {
                DataTable dt = new DataTable();
                string Import_FileName = path;
                string fileExtension = Path.GetExtension(Import_FileName);
                conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Import_FileName + ";" + "Extended Properties='Excel 12.0 Xml;HDR=YES;'";
                using (OleDbCommand comm = new OleDbCommand())
                {
                    comm.CommandText = "Select * from [" + sheetName + "$]";
                    comm.Connection = conn;
                    using (OleDbDataAdapter da = new OleDbDataAdapter())
                    {
                        da.SelectCommand = comm;
                        da.Fill(dt);
                        comm.Dispose();
                        return dt;
                    }
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Image Mapimage = gmap.ToImage();
            string ImagePath = Environment.CurrentDirectory + "MapImage.Png";
            Mapimage.Save(ImagePath);
            Mapimage.Dispose();
        }

        private void nud_zoom_ValueChanged(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void txtLong_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtLat_TextChanged(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void nud_maxzoom_ValueChanged(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void nud_minzoom_ValueChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {

        }

        private void txtCityName_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtCountryName_TextChanged(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void gmap_Load(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {

        }
    }
}
