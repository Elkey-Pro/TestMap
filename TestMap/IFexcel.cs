using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using DataTable = System.Data.DataTable;
using System.IO;
using Microsoft.Office.Core;
using XlChartType = Microsoft.Office.Interop.Excel.XlChartType;
using System.Drawing;

namespace TestMap
{
    public class IFexcel
    {
        public Excel.Application xlApp;
        public Excel.Workbook xlWorkBook;
        public Excel.Worksheet xlWorkSheet;
        public string Xlsx_output { get; set; }

        object misValue = System.Reflection.Missing.Value;
        object oFalse = false;
        object oTrue = true;

        public IFexcel(string ExcelPath)
        {
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(ExcelPath);
        }
            
          


        public class ExcelChart 
        { 
            int Left {  get; set; }
            int Top { get; set; }
            int Width { get; set; }
            int Height { get; set; }
            string ExportLocation { get; set; }
            string ExportName { get; set; }
            string ChartImgName { get;set; }
            string ChartTitle { get; set; }

            string ChartDataRage { get;set; }
            string XdataRage { get; set; }
            string YdataRage { get; set; }
            string SerieName { get;set; }
            Excel.XlChartType ChartType { get; set; }
            bool isLineNear { get; set; }
            Microsoft.Office.Core.MsoThemeColorIndex ChartColor { get;set; }
            Excel.XlRgbColor fontColor { get; set; }
            string DataSymbol { get; set; }

            public Excel.Worksheet WorkSheet;

            public ExcelChart(Excel.Worksheet XLWorkSheet)
            {
                WorkSheet = XLWorkSheet;
            }

            public void CreateChart(string outPutName, string outPutFolder )
            {
                Excel.ChartObjects chartObjects = (Excel.ChartObjects)WorkSheet.ChartObjects();
                Excel.ChartObject scatterChartObject = chartObjects.Add(Left,Top, Width, Height);
                Excel.Chart CustomChart = scatterChartObject.Chart;
                CustomChart.ChartType = ChartType;

                if (ChartType == XlChartType.xlXYScatter)
                {
                    CustomChart.SetSourceData(WorkSheet.Range[ChartDataRage]);


                    if (isLineNear)
                    {
                        Excel.Series Series1 = CustomChart.SeriesCollection().Item(1);

                        Excel.Trendlines trendlines = Series1.Trendlines();

                        Excel.Trendline trendline = trendlines.Add(Excel.XlTrendlineType.xlLinear);

                        trendline.Name = "";

                        trendline.Border.Color = Excel.XlRgbColor.rgbWhite;

                    }

                }
                else
                {
                    Excel.Range xValuesRange = WorkSheet.Range[XdataRage];
                    Excel.Range yValuesRange = WorkSheet.Range[YdataRage];
                    Excel.Series ySeries = CustomChart.SeriesCollection().NewSeries();
                    ySeries.Values = yValuesRange;
                    ySeries.Name = SerieName;
                    ySeries.XValues = xValuesRange;

                    if (!string.IsNullOrEmpty(DataSymbol))
                    {
                        yValuesRange.NumberFormat = "0" + DataSymbol;
                    }
                    if (isLineNear)
                    {
                        Excel.Series Series1 = CustomChart.SeriesCollection().Item(1);

                        Excel.Trendlines trendlines = Series1.Trendlines();

                        Excel.Trendline trendline = trendlines.Add(Excel.XlTrendlineType.xlLinear);

                        trendline.Name = "";

                        trendline.Border.Color = Excel.XlRgbColor.rgbDarkGreen;

                    }
                }
                CustomChart.HasTitle = true;

                CustomChart.ChartTitle.Text = ChartTitle;

                CustomChart.ChartTitle.Font.Color = Color.White;

                CustomChart.ChartStyle = 8;

                Excel.ChartArea chartArea = CustomChart.ChartArea;

                chartArea.Format.Fill.ForeColor.ObjectThemeColor = ChartColor;

                CustomChart.PlotArea.Format.Fill.ForeColor.ObjectThemeColor = ChartColor;


                CustomChart.ChartArea.Font.Color = fontColor;

                CustomChart.ApplyDataLabels(Excel.XlDataLabelsType.xlDataLabelsShowValue, false, true, false, false, false, true, true, false, false);

                // add DataSymbol
                if (!string.IsNullOrEmpty(DataSymbol))
                {
                    Excel.Series Series1 = CustomChart.SeriesCollection().Item(1);


                    foreach (Excel.DataLabel datalabel in Series1.DataLabels())
                    {
                        datalabel.NumberFormat = "0" + DataSymbol;
                    }
                }

                CustomChart.Export(outPutFolder + @"\" + outPutName, "PNG");

            }

        }

        public DataTable ReadDataExcel(string path, string SheetName)
        {
            using (OleDbConnection conn = new OleDbConnection())
            {
                DataTable dt = new DataTable();
                string Import_FileName = path;
                string fileExtension = Path.GetExtension(Import_FileName);
                conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Import_FileName + ";" + "Extended Properties='Excel 12.0 Xml;HDR=YES;'";
                using (OleDbCommand comm = new OleDbCommand())
                {
                    comm.CommandText = "Select * from [" + SheetName + "$]";
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

        public  void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                Console.WriteLine("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }


    }
}
