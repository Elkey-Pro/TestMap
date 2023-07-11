using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using GMap;
using GMap.NET.WindowsForms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using static RvAutoReport.IFexcel;
using Action = System.Action;
using DataTable = System.Data.DataTable;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using WORD = Microsoft.Office.Interop.Word;
using System.Threading;
using GMap.NET;
using Task = System.Threading.Tasks.Task;
using System.Xml.Linq;

namespace RvAutoReport
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }
        // for old script
        //static string Word_Report_Name = "RV Rental Tampa Report-v0.2.docx";


        static string xlsx_output = Environment.CurrentDirectory + @"\XLSX_OUTPUT";
        static string Logo_path = Environment.CurrentDirectory + @"\No-Backgrounds";
        string Word_report_template_file = Environment.CurrentDirectory + @"\template" + @"\" ;
        static string Word_output = Environment.CurrentDirectory + @"\WORD_OUTPUT";
        static string img_temp = Environment.CurrentDirectory + @"\temp_image";
        static object misValue = System.Reflection.Missing.Value;
        static object oFalse = false;
        static object oTrue = true;
        static string url_PriceNight = @"\priceNight.png";
        static string url_numberofRVs = @"\numberofRVs.png";
        static string url_lengthof5RV = @"\lengthof5RV.png";
        static string url_AvgDailyPriceNight = @"\AvgDailyPriceNight.png";
        static string url_AcAhTable = @"\AcAH.png";
        static string url_cityMap = @"\CityMap.png";
        static List<string> list_image_url = new List<string>();
        public System.Threading.CancellationTokenSource TokenSource;

        //  for new version
        //public string ExcelPath = @"";
        public DataTable Excel_data;
        //public List<string> list_image_url = new List<string>();
        DataTable DataForWordReport = new DataTable (){ TableName  = "DataForWOrdReport" };
        public readonly string DataXmlPath = "Data.xml";
        private void Form1_Load(object sender, EventArgs e)
        {
            txt_xlsxPath.Text = Environment.CurrentDirectory +  @"\XLSX_OUTPUT" ;
            txt_DocxOutPutPath.Text = Environment.CurrentDirectory + @"\WORD_OUTPUT";
            Word_report_template_file = Word_report_template_file + txt_docxTemp.Text;
            // load data.xml to datatable & gridview
            //DataForWordReport.Columns.Add("FindString");
            //DataForWordReport.Columns.Add("Row");
            //DataForWordReport.Columns.Add("Column");
            //DataForWordReport.Columns.Add("Type");
            //DataForWordReport.ReadXml(DataXmlPath);
            //dGVDataWFromXl.DataSource = DataForWordReport;

        }


        private static void KillWordAndExcelProcesses()
        {
            // Kill all running Word processes
            Process[] wordProcesses = Process.GetProcessesByName("WINWORD");
            foreach (Process process in wordProcesses)
            {
                process.Kill();
            }

            // Kill all running Excel processes
            Process[] excelProcesses = Process.GetProcessesByName("EXCEL");
            foreach (Process process in excelProcesses)
            {
                process.Kill();
            }
        }


        private void btn_MakeReport_Click_1(object sender, EventArgs e)
        {
            //KillWordAndExcelProcesses();

            //string[] files = Directory.GetFiles(txt_xlsxPath.Text, "*.xlsx");

            //foreach (string file in files)
            //{
            //    if (!file.Contains("~$")) // ignore the excel temp file
            //        ReadExcel(file);
            //}
        }

        public void Invoke_require(Action Action)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => Action()));
            }
            else
            {
                Action();
            }
        }

        private void WriteLog(string msg)
        {
            Invoke_require(() => txt_log.AppendText(msg + Environment.NewLine));
        }

        private static DataTable ConvertToDataTable(IEnumerable<dynamic> data)
        {
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("Type", typeof(string));
            dataTable.Columns.Add("Count", typeof(int));

            foreach (var item in data)
            {
                dataTable.Rows.Add(item.Type, item.Count);
            }

            return dataTable;
        }

        //private void ReadExcel(string ExcelPath)
        //{
        //    WriteLog("Start Read Excel");
        //    WriteLog("Reading File :" + ExcelPath);
        //    IFexcel ExcelApp = new IFexcel(ExcelPath);

        //    // read all data
        //    DataTable FullData = ExcelApp.ReadDataExcel(ExcelPath,"OriginalData" );

        //    //sort by REVIEWS, Current YTD Utilization desc
        //    FullData.DefaultView.Sort = "[REVIEWS] DESC";
        //    FullData.DefaultView.Sort += ",[Current YTD Utilization] DESC";
        //    FullData = FullData.DefaultView.ToTable();

        //    //int TotalRV = FullData.AsEnumerable()
        //    //                                  .Select(row => row.Field<string>("TYPE"))
        //    //                                  .Distinct()
        //    //                                  .Count();

        //    // datatable for number of RVs
        //    var groupedRows = from DataRow row in FullData.Rows
        //                      group row by row["TYPE"] into typeGroup
        //                      select new
        //                      {
        //                          Type = typeGroup.Key,
        //                          Count = typeGroup.Count()
        //                      };
        //    DataTable RvCountGroupByType = ConvertToDataTable(groupedRows);
        //    RvCountGroupByType.DefaultView.Sort = "[Type] ASC";
        //    RvCountGroupByType = RvCountGroupByType.DefaultView.ToTable();
        //    foreach (DataRow dtrow in RvCountGroupByType.Rows)
        //    {
        //        if (dtrow["TYPE"].ToString() == "A" || dtrow["TYPE"].ToString() == "B" || dtrow["TYPE"].ToString() == "C")
        //        {
        //            dtrow["TYPE"] = "CLASS " + dtrow["TYPE"].ToString();
        //        }
        //    }

        //    foreach(Excel.Worksheet worksheet in ExcelApp.xlWorkBook.Worksheets)
        //    {
        //        if (worksheet.Name.Contains("stat-"))
        //        {
        //            ExcelApp.xlWorkSheet = (Excel.Worksheet)ExcelApp.xlWorkBook.Worksheets.get_Item(worksheet.Index);

        //            // Extract City Name from Workbook Name
        //            List<string> cityNametemp = ExcelApp.xlWorkBook.Name.Split('-').ToList();

        //            string CityName = cityNametemp[1];

        //            list_image_url.Clear();


        //            // get the RvType
        //            string SheetName = ExcelApp.xlWorkSheet.Name.Replace("stat-", "");

        //            string RvType = string.Empty;

        //            if (SheetName == "A" || SheetName == "B" || SheetName == "C")
        //            {
        //                RvType = "Class " + SheetName;
        //            }
        //            else
        //            {
        //                RvType = SheetName;
        //            }

        //            // create Doc file name   =  city name - RvType
        //            string DocFileName = ExcelApp.xlWorkSheet.Name.Replace(CityName, CityName + "-" + RvType).Replace("xlsx", "docx");

        //            // read accompanying data set and sort 
        //            var filteredRow = FullData.AsEnumerable().Where(row => row.Field<string>("TYPE") == SheetName);

        //            DataTable Data = FullData.Clone();

        //            foreach (var row in filteredRow)
        //            {
        //                Data.ImportRow(row);
        //            }

        //            // sort data by REVIEWS ,Current YTD Utilization] DESC
        //            Data.DefaultView.Sort = "[REVIEWS] DESC";
        //            Data.DefaultView.Sort += ",[Current YTD Utilization] DESC";
        //            Data = Data.DefaultView.ToTable();

        //            // create table for word data report
        //            DataTable WordReport = new DataTable();
        //            WordReport.Columns.Add("FindString");
        //            WordReport.Columns.Add("ReplaceString");


        //            // start get data

        //            // get list_of top5 maker
        //            List<string> Top5MakerandModel = Data.AsEnumerable().Take(5)
        //                                                    .Select(row => row.Field<string>("MAKE") + " - Length : " + row.Field<Double>("LENGTH").ToString())
        //                                                    .Distinct()
        //                                                    .ToList();


        //            // read data from excel file
        //            foreach (DataRow datarow in DataForWordReport.Rows)
        //            {
        //                if (!string.IsNullOrEmpty(datarow.Field<string>("FindString")))
        //                {
        //                    var celldata = ExcelApp.xlWorkSheet.Cells[int.Parse(datarow.Field<string>("Row")), int.Parse(datarow.Field<string>("Column"))].Value;

        //                    if (!string.IsNullOrEmpty(datarow.Field<string>("Type")))
        //                    {
        //                        celldata = (celldata * 100) + "%";
        //                    }

        //                    WordReport.Rows.Add(datarow.Field<string>("FindString"), celldata);
        //                }

        //            }

        //            // add city Name
        //            WordReport.Rows.Add("<CITY_LIST>", CityName);

        //            // add  report date
        //            WordReport.Rows.Add("<REPORT_DATE>", DateTime.Now.ToString("MMMM dd,yyyy"));

        //            // add RV type
        //            WordReport.Rows.Add("<RV_TYPE>", RvType);

        //            // add total RV in report area
        //            WordReport.Rows.Add("<TOTAL_RV>", FullData.Rows.Count);

        //            // add total RV type
        //            WordReport.Rows.Add("<TOTAL_RV_TYPE>", Data.Rows.Count);

        //            // Count Total Rv
        //            #region totalRVcount
        //            //write header to cell
        //            ExcelApp.xlWorkSheet.Cells[1, 28].Value = "TYPE";
        //            ExcelApp.xlWorkSheet.Cells[1, 29].Value = "COUNT";

        //            // insert totalRVcount data set to excel

        //            for (int i = 0; i < RvCountGroupByType.Rows.Count; i++)
        //            {

        //                ExcelApp.xlWorkSheet.Cells[i + 2, 28].Value = RvCountGroupByType.Rows[i].Field<string>(0);
        //                ExcelApp.xlWorkSheet.Cells[i + 2, 29].Value = RvCountGroupByType.Rows[i].Field<int>(1).ToString();

        //            }

        //            // insert totalRVcount chart to excel

        //            string chartImgName = "numberofRVs.png";
        //            string FolderPath = Environment.CurrentDirectory + @"\temp_image\" + RvType;
        //            string Start = "AB";
        //            string End = "AC";

        //            if (!Directory.Exists(FolderPath))
        //            {
        //                Directory.CreateDirectory(FolderPath);
        //            }


        //            ExcelChart totalRVcountChart = new ExcelChart()
        //            {
        //                WorkSheet = ExcelApp.xlWorkSheet,
        //                ChartColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorLight1,
        //                ChartTitle = "TOTAL RVs Count",
        //                ChartType = XlChartType.xlBarClustered,
        //                fontColor = XlRgbColor.rgbBlack,
        //                DataSymbol = "",
        //                isLineNear = false,
        //                SerieName = "RV Count",
        //                Left = 5,
        //                Top = 650,
        //                Width = 500,
        //                Height = 300,
        //                XdataRage = "$" + Start + "$2:$" + Start + "$" + (RvCountGroupByType.Rows.Count + 1),
        //                YdataRage = "$" + End + "$2:$" + End + "$" + (RvCountGroupByType.Rows.Count + 1),
        //                ChartDataRage = Start + ":" + End
        //            };
        //            totalRVcountChart.CreateChart(chartImgName, FolderPath);

        //            //ExcelApp.CreateChart(xlWorkSheet, "AB", "AC", (1 + RvCountGroupByType.Rows.Count), "TOTAL RVs Count"
        //            //    , "RV Count", XlChartType.xlBarClustered, 5, 650, 500, 300, RvType, url_numberofRVs, Office.MsoThemeColorIndex.msoThemeColorLight1, Color.Black, Excel.XlRgbColor.rgbBlack);

        //            #endregion     

        //        }                

        //    }
        //    ExcelApp.xlWorkBook.SaveAs(ExcelPath, Excel.XlFileFormat.xlOpenXMLWorkbook, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlNoChange, misValue, misValue, misValue, misValue, misValue);
        //    ExcelApp.xlWorkBook.Close(true, misValue, misValue);
        //    ExcelApp.xlApp.Quit();
        //    ExcelApp.releaseObject();
        //}

        private DataTable GetDataTableFromDataGridView(DataGridView dataGridView)
        {
            DataTable dataTable = new DataTable();

            // Create columns in the DataTable
            foreach (DataGridViewColumn dataGridViewColumn in dataGridView.Columns)
            {
                dataTable.Columns.Add(dataGridViewColumn.Name, typeof(string));
            }

            // Populate rows in the DataTable
            foreach (DataGridViewRow dataGridViewRow in dataGridView.Rows)
            {
                DataRow dataRow = dataTable.NewRow();

                foreach (DataGridViewCell dataGridViewCell in dataGridViewRow.Cells)
                {
                    int columnIndex = dataGridViewCell.ColumnIndex;
                    dataRow[columnIndex] = dataGridViewCell.Value?.ToString() ?? string.Empty;
                }

                dataTable.Rows.Add(dataRow);
            }

            return dataTable;
        }


        private void button3_Click(object sender, EventArgs e)
        {
       
            //save Gridview to datatable
            DataForWordReport = GetDataTableFromDataGridView(dGVDataWFromXl);
            //save to file
            File.Delete(DataXmlPath);
            DataForWordReport.TableName = "DataForWOrdReport";
            DataForWordReport.WriteXml(DataXmlPath);
        }

        #region Copy From preview Script
        private void button4_Click(object sender, EventArgs e)
        {
            txt_log.Clear();
            WriteLog("Kill Excel , Word instance");
            KillWordAndExcelProcesses();

            string[] files = Directory.GetFiles(xlsx_output, "*.xlsx");

            WriteLog("Start Loop All file xlsx in folder");

            TokenSource = new System.Threading.CancellationTokenSource();
            CancellationToken token = TokenSource.Token;
            Task.Factory.StartNew(() =>

            {
                try
                {
                    foreach (string file in files)
                    {
                        if (!file.Contains("~$")) // ignore the excel temp file
                        {
                            ReadExcel(file);

                            if (token.IsCancellationRequested)
                            {
                                token.ThrowIfCancellationRequested();
                            }
                        }
                        WriteLog("DONE!!!!!!!!!");
                            
                    }
               }
                catch (OperationCanceledException) {  }
                catch (ObjectDisposedException) {  }
                catch (Exception ex)
                {
                    WriteLog(DateTime.Now + " " + ex.ToString());
                }

            }, token, TaskCreationOptions.LongRunning, TaskScheduler.Default);
         

        }

        private void CreateCityMap(DataTable RvData)
        {

            if (RvData.Rows.Count > 0)
            {
                this.Invoke(new MethodInvoker(delegate ()
                {
                    gmap.MaxZoom = 18;
                    gmap.MinZoom = 2;
                    gmap.Zoom = 8;
                    gmap.MapProvider = GMap.NET.MapProviders.GoogleMapProvider.Instance;
                    GMap.NET.GMaps.Instance.Mode = GMap.NET.AccessMode.ServerOnly;


                    DataTable LatLongData = new DataTable();
                    LatLongData.Columns.Add("LAT", typeof(double));
                    LatLongData.Columns.Add("LONG", typeof(double));

                    foreach (DataRow dtrow in RvData.Rows)
                    {
                        double Lat = (double)dtrow["LATITUDE"];
                        double Long = (double)dtrow["LONGITUDE"];
                        LatLongData.Rows.Add(Lat, Long);
                    }

                    // avg lat long
                    double AvgLat = LatLongData.AsEnumerable().Average(row => row.Field<double>("LAT"));
                    double AvgLong = LatLongData.AsEnumerable().Average(row => row.Field<double>("LONG"));


                    int numberOfPoints = 100;

                    // convert from mile to meter
                    double radiusInMeters = Double.Parse(txt_circleDiameter.Text) * 1609.344;

                    List<GMap.NET.PointLatLng> circlePoints = new List<GMap.NET.PointLatLng>();
                    double angle = 2 * Math.PI / numberOfPoints;
                    for (int i = 0; i < numberOfPoints; i++)
                    {
                        double lat = AvgLat + radiusInMeters / 111320d * Math.Sin(i * angle);
                        double lng = AvgLong + radiusInMeters / (111320d * Math.Cos(AvgLat * Math.PI / 180)) * Math.Cos(i * angle);
                        circlePoints.Add(new GMap.NET.PointLatLng(lat, lng));
                    }

                    GMap.NET.WindowsForms.GMapOverlay Overlay = new GMap.NET.WindowsForms.GMapOverlay("Overlay");
                    GMap.NET.WindowsForms.GMapPolygon circle = new GMap.NET.WindowsForms.GMapPolygon(circlePoints, "circle");

                    circle.Fill = new SolidBrush(Color.FromArgb(30, Color.Blue));  // Fill color
                    circle.Stroke = new Pen(Color.Blue, 1);

                    Overlay.Polygons.Add(circle);

                    gmap.Overlays.Add(Overlay);

                    // Create a GMapMarker for the center

                    // Add the marker to the overlay              


                    // add all location to overlay
                    foreach (DataRow dtrow in LatLongData.Rows)
                    {
                        GMap.NET.WindowsForms.GMapOverlay markers = new GMap.NET.WindowsForms.GMapOverlay("markers");
                        //GMap.NET.WindowsForms.GMapMarker marker = new GMap.NET.WindowsForms.Markers.GMarkerGoogle(
                        //                                                new GMap.NET.PointLatLng(dtrow.Field<double>("lat"), dtrow.Field<double>("long")),
                        //                                                            GMap.NET.WindowsForms.Markers.GMarkerGoogleType.red_small);
                        //Overlay = new GMap.NET.WindowsForms.GMapOverlay("markers");
                        GMap.NET.WindowsForms.GMapMarker marker = new GMap.NET.WindowsForms.Markers.GMarkerGoogle(
                                                                        new GMap.NET.PointLatLng(dtrow.Field<double>("lat"), dtrow.Field<double>("long")),
                                                                                    GMap.NET.WindowsForms.Markers.GMarkerGoogleType.red_small);
                        markers.Markers.Add(marker);
                        gmap.Overlays.Add(markers);

                    }            

                    


                    gmap.Position = new GMap.NET.PointLatLng(AvgLat, AvgLong);
                }
            ));
            }

            



        }

        private static void Insert_chart(Excel.Worksheet xlWorkSheet, string chart_range_start, string chart_range_end, int total_line, string chart_name, string serie_name, Excel.XlChartType chartType
           , double left, double top, double width, double high, string FolderPath, string exported_image_Name
           , Microsoft.Office.Core.MsoThemeColorIndex ChartColor, Color TitleCOlor, Excel.XlRgbColor fontColor, string Currency = "",
           bool isAddLinear = false)
        {

            string xDataRange = "$" + chart_range_start + "$2:$" + chart_range_start + "$" + total_line; // Replace with the range of your X values
            string yDataRange = "$" + chart_range_end + "$2:$" + chart_range_end + "$" + total_line; // Replace with the range of your Y values

            Excel.ChartObjects chartObjects = (Excel.ChartObjects)xlWorkSheet.ChartObjects();
            Excel.ChartObject scatterChartObject = chartObjects.Add(5, 350, 800, 300);
            Excel.Chart CustomChart = scatterChartObject.Chart;

            CustomChart.ChartType = chartType;

            if (chartType == XlChartType.xlXYScatter)
            {
                CustomChart.SetSourceData(xlWorkSheet.Range[chart_range_start + ":" + chart_range_end]);


                if (isAddLinear)
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
                Excel.Range xValuesRange = xlWorkSheet.Range[xDataRange];
                //Excel.Series xSeries = CustomChart.SeriesCollection().NewSeries();
                //xSeries.XValues = xValuesRange;


                Excel.Range yValuesRange = xlWorkSheet.Range[yDataRange];
                Excel.Series ySeries = CustomChart.SeriesCollection().NewSeries();
                ySeries.Values = yValuesRange;
                ySeries.Name = serie_name;
                ySeries.XValues = xValuesRange;

                if (!string.IsNullOrEmpty(Currency))
                {
                    yValuesRange.NumberFormat = "0" + Currency;
                }

                if (isAddLinear)
                {
                    Excel.Series Series1 = CustomChart.SeriesCollection().Item(1);

                    Excel.Trendlines trendlines = Series1.Trendlines();

                    Excel.Trendline trendline = trendlines.Add(Excel.XlTrendlineType.xlLinear);

                    trendline.Name = "";

                    trendline.Border.Color = Excel.XlRgbColor.rgbDarkGreen;

                }


            }


            CustomChart.HasTitle = true;

            CustomChart.ChartTitle.Text = chart_name;

            CustomChart.ChartTitle.Font.Color = Color.White;

            CustomChart.ChartStyle = 8;

            Excel.ChartArea chartArea = CustomChart.ChartArea;

            chartArea.Format.Fill.ForeColor.ObjectThemeColor = ChartColor;

            CustomChart.PlotArea.Format.Fill.ForeColor.ObjectThemeColor = ChartColor;


            CustomChart.ChartArea.Font.Color = fontColor;

            CustomChart.ApplyDataLabels(Excel.XlDataLabelsType.xlDataLabelsShowValue, false, true, false, false, false, true, true, false, false);

            // add currency
            if (!string.IsNullOrEmpty(Currency))
            {
                Excel.Series Series1 = CustomChart.SeriesCollection().Item(1);


                foreach (Excel.DataLabel datalabel in Series1.DataLabels())
                {
                    datalabel.NumberFormat = "0" + Currency;
                }
            }

            CustomChart.Export(FolderPath + @"\" + exported_image_Name, "PNG");

        }

        private  void ReadExcel(string ExcelPath)
        { 
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;


            xlApp = new Excel.Application();
            xlApp.DisplayAlerts = false;
            // open the file
            xlWorkBook = xlApp.Workbooks.Open(ExcelPath);

            WriteLog("Read Data from : " + xlWorkBook.Name);

            // read all data
            DataTable FullData = ReadExcelFile("OriginalData", ExcelPath);

            //sort by Current YTD Utilization desc
            FullData.DefaultView.Sort = "[Current YTD Utilization] DESC";
            FullData = FullData.DefaultView.ToTable();

            // Total RV

            int TotalRV = FullData.AsEnumerable()
                                                .Select(row => row.Field<string>("TYPE"))
                                                .Distinct()
                                                .Count();

            // datatable for number of RVs
            var groupedRows = from DataRow row in FullData.Rows
                              group row by row["TYPE"] into typeGroup
                              select new
                              {
                                  Type = typeGroup.Key,
                                  Count = typeGroup.Count()
                              };
            DataTable RvCountGroupByType = ConvertToDataTable(groupedRows);
            RvCountGroupByType.DefaultView.Sort = "[Type] ASC";
            RvCountGroupByType = RvCountGroupByType.DefaultView.ToTable();
            foreach (DataRow dtrow in RvCountGroupByType.Rows)
            {
                if (dtrow["TYPE"].ToString() == "A" || dtrow["TYPE"].ToString() == "B" || dtrow["TYPE"].ToString() == "C")
                {
                    dtrow["TYPE"] = "CLASS " + dtrow["TYPE"].ToString();
                }
            }


            // start loop all excel stat- sheet
            foreach (Excel.Worksheet shittt in xlWorkBook.Worksheets)
            {
                if (shittt.Name.Contains("stat-"))
                {
                    WriteLog("Read Data from sheet: " + shittt.Name);

                    List<string> cityNametemp = xlWorkBook.Name.Split('-').ToList();

                    string CityName = cityNametemp[1];

                    list_image_url.Clear();

                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(shittt.Index);


                    string SheetName = shittt.Name.Replace("stat-", "");

                    
                    // get the RvType
                    string RvType = string.Empty;

                    if (SheetName == "A" || SheetName == "B" || SheetName == "C")
                    {
                        RvType = "Class " + SheetName;
                    }
                    else
                    {
                        RvType = SheetName;
                    }
                    WriteLog("RvType : "+ RvType);

                    // create folder for template image file
                    string FolderPath = img_temp + @"\" + RvType;

                    if (!Directory.Exists(FolderPath))
                    {
                        Directory.CreateDirectory(FolderPath);
                    }
                    WriteLog("Image Folder Path : " + FolderPath);

                    // create Doc file name   =  city name - RvType
                    string DocFileName = xlWorkBook.Name.Replace(CityName, CityName + "-" + RvType).Replace("xlsx", "docx");

                    // read accompanying data set and sort 
                    var filteredRow = FullData.AsEnumerable().Where(row => row.Field<string>("TYPE") == SheetName);

                    DataTable Data = FullData.Clone();

                    foreach (var row in filteredRow)
                    {
                        Data.ImportRow(row);
                    }

                    // sort by REVIEWS
                    Data.DefaultView.Sort += "[Current YTD Utilization] DESC";
                    Data = Data.DefaultView.ToTable();
                    WriteLog("Sorting"  + RvType + " Data");                

                    // check if lat & long data is empty or not
                    DataTable RvData = new DataTable();
                    var checkifemptydata = Data.AsEnumerable()
                    .Where(row => !row.IsNull("LATITUDE") && !row.IsNull("LONGITUDE"));

                    if(checkifemptydata.Any())
                    {
                        RvData = checkifemptydata.CopyToDataTable();
                    }


                    if (RvData.Rows.Count > 0)
                    {
                        WriteLog("Creating Map");

                        // create map
                        CreateCityMap(RvData);

                        // wait for the map render completely
                        Thread.Sleep(2000);

                        // export map image
                        this.Invoke(new MethodInvoker(delegate ()
                        {
                            Image Mapimage = gmap.ToImage();
                            string ImagePath = FolderPath + url_cityMap;
                            Mapimage.Save(ImagePath);
                            Mapimage.Dispose();

                            foreach (GMapOverlay overlay in gmap.Overlays.ToList())
                            {
                                gmap.Overlays.Remove(overlay);
                            }
                            WriteLog("Map image exported");

                        }));
                       
                    }
                    else
                    {
                        WriteLog("missing lat & long data");
                    }
                   
                    DataTable DataForWordReport = new DataTable("DataForWOrdReport");
                    DataForWordReport.Columns.Add("FindString");
                    DataForWordReport.Columns.Add("ReplaceString");

                    WriteLog(" get List top 5 Maker");
                    // get list_of top5 maker
                    //List<string> Top5MakerandModel = Data.AsEnumerable().Take(5)
                    //                                        .Select(row => row.Field<string>("MAKE") + " - Length : " + row.Field<Double>("LENGTH").ToString())
                    //                                        .Distinct()
                                                            //.ToList();
                    List<string> Top5MakerandModel = Data.AsEnumerable()
                                        .Select((row) =>  row.Field<string>("MAKE") )
                                        .Distinct().Take(5)
                                        .ToList();
                    for(int i  = 0; i < Top5MakerandModel.Count; i++ )
                    {
                        Top5MakerandModel[i] = (i +1).ToString() +". " + Top5MakerandModel[i];
                    }

                    DataRow dtrow = DataForWordReport.NewRow();

                    WriteLog("Adding Data from Stat sheet to data table");

                    // add User Name
                    DataForWordReport.Rows.Add("<USER_NAME>", txt_UserName.Text);

                    // add city Name
                    DataForWordReport.Rows.Add("<CITY_LIST>", CityName);

                    // add  report date
                    DataForWordReport.Rows.Add("<REPORT_DATE>", DateTime.Now.ToString("MMMM dd,yyyy"));

                    // add RV type
                    DataForWordReport.Rows.Add("<RV_TYPE>", RvType);

                    // add total RV in report area
                    DataForWordReport.Rows.Add("<TOTAL_RV>", FullData.Rows.Count);

                    // add total RV type
                    DataForWordReport.Rows.Add("<TOTAL_RV_TYPE>", Data.Rows.Count);



                    #region Load data from excel

                    xlWorkSheet.Activate();

                    // add Average Age of all RVs
                    DataForWordReport.Rows.Add("<ALL_AVG_RV_AGE>", xlWorkSheet.Cells[4, 2].Value);

                    //add Average Age of the Top 5 Class 
                    DataForWordReport.Rows.Add("<TOP5_AVG_RV_AGE>", xlWorkSheet.Cells[5, 2].Value);

                    //add Average Age of the Top 25 Class 
                    DataForWordReport.Rows.Add("<TOP25_AVG_RV_AGE>", xlWorkSheet.Cells[6, 2].Value);

                    //add Average length of the Top 5 Class 
                    DataForWordReport.Rows.Add("<TOP5_AVG_RV_LENGTH>", xlWorkSheet.Cells[7, 2].Value);

                    //Average Utilization In Season (May 1 to Oct 31) Top 5

                    DataForWordReport.Rows.Add("<TOP5_AVG_UTI_IN_SEASON>", (xlWorkSheet.Cells[15, 2].Value * 100) + "%");

                    //Average Utilization Off Season (Nov1 to Apr30) Top 5
                    DataForWordReport.Rows.Add("<TOP5_AVG_UTI_OFF_SEASON>", (xlWorkSheet.Cells[16, 2].Value * 100) + "%");

                    //Average Utilization In Season (May 1 to Oct 31) Top 25
                    DataForWordReport.Rows.Add("<TOP25_AVG_UTI_IN_SEASON>", (xlWorkSheet.Cells[17, 2].Value * 100) + "%");

                    //Average Utilization Off Season (Nov1 to Apr30) Top 25
                    DataForWordReport.Rows.Add("<TOP25_AVG_UTI_OFF_SEASON>", (xlWorkSheet.Cells[18, 2].Value * 100) + "%");

                    //Average Utilization 2021 Top 5
                    DataForWordReport.Rows.Add("<TOP5_AVG_UTI_2021>", (xlWorkSheet.Cells[11, 2].Value * 100) + "%");

                    //Average Utilization 2022 Top 5
                    DataForWordReport.Rows.Add("<TOP5_AVG_UTI_2022>", (xlWorkSheet.Cells[12, 2].Value * 100) + "%");

                    //Average Utilization of top 5 YTD
                    DataForWordReport.Rows.Add("<TOP5_AVG_UTI_YTD>", (xlWorkSheet.Cells[13, 2].Value * 100) + "%");

                    //Average Utilization of Top 25 YTD
                    DataForWordReport.Rows.Add("<TOP25_AVG_UTI_YTD>", (xlWorkSheet.Cells[14, 2].Value * 100) + "%");

                    //Future utilization 30 Days
                    DataForWordReport.Rows.Add("<FUTURE_UTI_30>", (xlWorkSheet.Cells[19, 2].Value * 100) + "%");

                    //Future utilization 60 Days
                    DataForWordReport.Rows.Add("<FUTURE_UTI_60>", (xlWorkSheet.Cells[20, 2].Value * 100) + "%");

                    //Future utilization 90 Days
                    DataForWordReport.Rows.Add("<FUTURE_UTI_90>", (xlWorkSheet.Cells[21, 2].Value * 100) + "%");

                    //  Potential Annual
                    DataForWordReport.Rows.Add("<POTENTIAL_ANNUAL>", xlWorkSheet.Cells[22, 2].Value);

                    //  Average nightly price of top 5 RVs

                    //double AvgTop5NightlyPrice = xlWorkSheet.Cells[9, 2].Value;

                    Data.DefaultView.Sort = "[UTIL 2022] DESC";
                    DataTable Data2 = Data.DefaultView.ToTable();

                    double AvgTop5NightlyPrice = Math.Round( Data2.AsEnumerable().Take(5).Select(row => row.Field<double>("PRICE/NIGHT")).Average(),2) ;


                    DataForWordReport.Rows.Add("<AVG_NP_TOP5>", AvgTop5NightlyPrice);

                    //  Average nightly price calculation for last table

                    DataForWordReport.Rows.Add("<AVG_NP_TOP5_025>", Math.Round(AvgTop5NightlyPrice * ( 0.25 * 365 ) ));

                    DataForWordReport.Rows.Add("<AVG_NP_TOP5_026>", Math.Round(AvgTop5NightlyPrice *  (0.26 * 365) ));

                    DataForWordReport.Rows.Add("<AVG_NP_TOP5_050>", Math.Round(AvgTop5NightlyPrice * ( 0.50 * 365 ) ));

                    DataForWordReport.Rows.Add("<AVG_NP_TOP5_051>", Math.Round(AvgTop5NightlyPrice * ( 0.51 * 365 )));

                    DataForWordReport.Rows.Add("<AVG_NP_TOP5_075>", Math.Round(AvgTop5NightlyPrice * ( 0.75 * 365 )));

                    DataForWordReport.Rows.Add("<AVG_NP_TOP5_076>", Math.Round(AvgTop5NightlyPrice * (0.76 * 365)));

                    DataForWordReport.Rows.Add("<AVG_NP_TOP5_100>", Math.Round(AvgTop5NightlyPrice * 365 ) );



                    #endregion


                    DataForWordReport.WriteXml("data.xml");


                    WriteLog("Start Create Chart");
                    #region totalRVcount
                    //write header to cell
                    xlWorkSheet.Cells[1, 28].Value = "TYPE";
                    xlWorkSheet.Cells[1, 29].Value = "COUNT";

                    // insert totalRVcount data set to excel

                    for (int i = 0; i < RvCountGroupByType.Rows.Count; i++)
                    {

                        xlWorkSheet.Cells[i + 2, 28].Value = RvCountGroupByType.Rows[i].Field<string>(0);
                        xlWorkSheet.Cells[i + 2, 29].Value = RvCountGroupByType.Rows[i].Field<int>(1).ToString();

                    }

                    // insert totalRVcount chart to excel

                    Insert_chart(xlWorkSheet, "AB", "AC", (1 + RvCountGroupByType.Rows.Count), "TOTAL RVs Count"
                        , "RV Count", XlChartType.xlBarClustered, 5, 650, 500, 300, FolderPath, url_numberofRVs, Office.MsoThemeColorIndex.msoThemeColorLight1, Color.Black, Excel.XlRgbColor.rgbBlack);

                    WriteLog("Created Chart for totalRVcount ");

                    #endregion


                    #region top5 Rv length
                    // avg top5 Rv length data set
                    var temp = Data.AsEnumerable().Where(row => row.Field<double>("LENGTH") > 0).Select(row => new
                    {
                        MAKE = row.Field<string>("MAKE"),
                        LENGTH = row.Field<double>("LENGTH")
                    }).Take(5);
                    DataTable Top5_length = new DataTable();
                    Top5_length.Columns.Add("MAKE", typeof(string));
                    Top5_length.Columns.Add("LENGTH", typeof(int));

                    foreach (var item in temp)
                    {
                        Top5_length.Rows.Add(item.MAKE, item.LENGTH);
                    }
                    // insert header
                    xlWorkSheet.Cells[1, 30].Value = "MAKE";
                    xlWorkSheet.Cells[1, 31].Value = "LENGTH";

                    // insert  avg top5 Rv length data set to excel

                    for (int i = 0; i < Top5_length.Rows.Count; i++)
                    {

                        xlWorkSheet.Cells[i + 2, 30].Value = Top5_length.Rows[i].Field<string>(0);
                        xlWorkSheet.Cells[i + 2, 31].Value = Top5_length.Rows[i].Field<int>(1);

                    }

                    // insert avg top5 Rv length chart to excel
                    WriteLog("Created Chart for  avg top5 Rv length ");

                    Insert_chart(xlWorkSheet, "AD", "AE", (1 + Top5_length.Rows.Count), "Top5 Rv length"
                    , "RV LENGTH", XlChartType.xlColumnClustered, 5, 950, 500, 300, FolderPath, url_lengthof5RV, Office.MsoThemeColorIndex.msoThemeColorLight1, Color.Black, Excel.XlRgbColor.rgbBlack, "", true);
                    #endregion


                    #region Average Daily Price/Nigh

                    // data set for Average Daily Price/Nigh
                    DataTable avg_daily_pricenight = new DataTable();
                    avg_daily_pricenight.Columns.Add("TYPE", typeof(string));
                    avg_daily_pricenight.Columns.Add("VALUE", typeof(double));

                    //string type = FullData.Columns[8].DataType.ToString();

                    decimal avgtop5 = Data.AsEnumerable().Take(5).Average(row => row.Field<decimal>("PRICE/NIGHT"));

                    decimal avgtop25 = Data.AsEnumerable().Take(25).Average(row => row.Field<decimal>("PRICE/NIGHT"));
                    decimal avgAll = Data.AsEnumerable().Average(row => row.Field<decimal>("PRICE/NIGHT"));

                    avg_daily_pricenight.Rows.Add("TOP 5", (int)Math.Round(avgtop5));
                    avg_daily_pricenight.Rows.Add("TOP 25", (int)Math.Round(avgtop25));
                    avg_daily_pricenight.Rows.Add("All", (int)Math.Round(avgAll));

                    // insert Average Daily Price/Night to excel

                    // insert header
                    xlWorkSheet.Cells[1, 32].Value = "TYPE";
                    xlWorkSheet.Cells[1, 33].Value = "VALUE";
                    // insert data
                    for (int i = 0; i < avg_daily_pricenight.Rows.Count; i++)
                    {

                        xlWorkSheet.Cells[i + 2, 32].Value = avg_daily_pricenight.Rows[i].Field<string>(0);
                        xlWorkSheet.Cells[i + 2, 33].Value = avg_daily_pricenight.Rows[i].Field<double>(1);

                    }
                    Insert_chart(xlWorkSheet, "AF", "AG", (1 + avg_daily_pricenight.Rows.Count), "Average Price / Night"
                     , "Price/Night", XlChartType.xlBarClustered, 5, 1250, 500, 300, FolderPath, url_AvgDailyPriceNight, Office.MsoThemeColorIndex.msoThemeColorLight1, Color.Black, Excel.XlRgbColor.rgbBlack, "$");
                   
                    WriteLog("Created Chart for  Average Daily Price/Night  ");
                    #endregion

                    #region ACtoAH chart

                    // create datasource

                    double Top25PETFRIENDELY = Data.AsEnumerable().Take(25).Sum(row => row.Field<double>("PET FRIENDELY")) / 25;
                    double Top25TAILGATEFRIENDELY = Data.AsEnumerable().Take(25).Sum(row => row.Field<double>("TAILGATE FRIENDELY")) / 25;
                    double Top25SMOKINGALLOWED = Data.AsEnumerable().Take(25).Sum(row => row.Field<double>("SMOKING ALLOWED")) / 25;
                    double Top25FESTIVALFRIENDLY = Data.AsEnumerable().Take(25).Sum(row => row.Field<double>("FESTIVAL FRIENDLY")) / 25;
                    double Top25GENERATOR = Data.AsEnumerable().Take(25).Sum(row => row.Field<double>("GENERATOR")) / 25;

                    DataTable AcAHTable = new DataTable();
                    AcAHTable.Columns.Add("TYPE");
                    AcAHTable.Columns.Add("VALUE");
                    AcAHTable.Rows.Add("PET FRIENDELY", Top25PETFRIENDELY);
                    AcAHTable.Rows.Add("TAILGATE FRIENDELY", Top25TAILGATEFRIENDELY);
                    AcAHTable.Rows.Add("SMOKING ALLOWED", Top25SMOKINGALLOWED);
                    AcAHTable.Rows.Add("FESTIVAL FRIENDLY", Top25FESTIVALFRIENDLY);
                    AcAHTable.Rows.Add("GENERATOR", Top25GENERATOR);

                    // insert data to excel

                    xlWorkSheet.Cells[1, 34].Value = "TYPE";
                    xlWorkSheet.Cells[1, 35].Value = "VALUE";

                    for (int i = 0; i < AcAHTable.Rows.Count; i++)
                    {

                        xlWorkSheet.Cells[i + 2, 34].Value = AcAHTable.Rows[i].Field<string>(0);
                        xlWorkSheet.Cells[i + 2, 35].Value = AcAHTable.Rows[i].Field<string>(1);

                    }
                    WriteLog("Created Chart for  Attributes rate ");

                    Insert_chart(xlWorkSheet, "AH", "AI", (1 + AcAHTable.Rows.Count), "Attributes"
                     , "Rate", XlChartType.xlBarClustered, 5, 1250, 500, 300, FolderPath, url_AcAhTable, Office.MsoThemeColorIndex.msoThemeColorLight1, Color.Black, Excel.XlRgbColor.rgbBlack, "%");




                    #endregion

                    #region price/night scratchart

                    // cal top 5 price /night by type

                    DataRow[] filteredRows = Data.AsEnumerable().Where(rn => rn.Field<double>("YEAR") > 0).Take(5).ToArray();

                    DataTable ds_top5_pricenight = new DataTable();
                    ds_top5_pricenight.Columns.Add("PRICE/NIGHT", typeof(string));
                    ds_top5_pricenight.Columns.Add("YEAR", typeof(string));
                    ds_top5_pricenight.Columns.Add("NAME", typeof(string));

                    foreach (DataRow row in filteredRows)
                    {
                        ds_top5_pricenight.Rows.Add(row["NAME"], row["YEAR"], row["PRICE/NIGHT"]);
                    }

                    // read all sheet calculated data
                    //DataTable CurremtSheetCalculcatedData = new DataTable();
                    //CurremtSheetCalculcatedData.Columns.Add("TYPE");
                    //CurremtSheetCalculcatedData.Columns.Add("VALUE");

                    //for (int i = 1; i <= 22; i++)
                    //{
                    //    DataRow dtrow1 = CurremtSheetCalculcatedData.NewRow();
                    //    dtrow1["TYPE"] = xlWorkSheet.Cells[i, 1].Value;
                    //    dtrow1["VALUE"] = xlWorkSheet.Cells[i, 2].Value;
                    //    CurremtSheetCalculcatedData.Rows.Add(dtrow1);
                    //}

                    //write header to cell
                    xlWorkSheet.Cells[1, 25].Value = "NAME";
                    xlWorkSheet.Cells[1, 26].Value = "YEAR";
                    xlWorkSheet.Cells[1, 27].Value = "PRICE/NIGHT";

                    // insert price/night data set to excel

                    for (int i = 0; i < ds_top5_pricenight.Rows.Count; i++)
                    {
                        for (int j = 0; j < 3; j++)
                        {
                            xlWorkSheet.Cells[i + 2, j + 25].Value = ds_top5_pricenight.Rows[i].Field<string>(j);
                        }
                    }
                    // insert price/night chart to excel

                    Insert_chart(xlWorkSheet, "Z", "AA", (1 + ds_top5_pricenight.Rows.Count), "Top 5 RV PRICE / NIGHT"
                        , "PRICE/NIGHT", XlChartType.xlXYScatter, 5, 350, 800, 300, FolderPath, url_PriceNight, Office.MsoThemeColorIndex.msoThemeColorAccent1, Color.White, Excel.XlRgbColor.rgbWhite, "", true);

                    WriteLog("Created Chart for  Top 5 RV PRICE / NIGHT ");

                    #endregion


                    WriteLog("Inserting Logo to excel sheet ");
                    // start insert logo
                    string cell = string.Empty;
                    //load all the available logo
                    for (int i = 1; i <= 5; i++)
                    {
                        cell = xlWorkSheet.Cells[i, 5].Value;
                        if (!string.IsNullOrEmpty(cell))
                        {
                            // add image url 
                            list_image_url.Add(cell);

                            // clear values
                            xlWorkSheet.Cells[i, 5].Value = "";
                        }
                    }

                    // insert logo to sheet
                    if (list_image_url.Count > 0)
                    {
                        int x = 500;
                        int y = 0;


                        int width = 0;
                        int height = 80;
                        foreach (var item in list_image_url)
                        {
                            System.Drawing.Image Image = System.Drawing.Image.FromFile(item);
                            // get ogirinal image ratio
                            double ratio = (double)Image.Width / Image.Height;

                            // change the width following original ratio
                            width = (int)(height * ratio);
                            //add
                            xlWorkSheet.Shapes.AddPicture(item, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, x, y, width, height);
                            y = y + height;
                        }
                    }

                    // generate docx report file
                    WriteLog("Generating Word report ");

                    Generate_Word_Report(RvType, xlWorkBook.Name.Replace(".xlsx", ""), DataForWordReport, Top5MakerandModel, DocFileName);

                    releaseObject(xlWorkSheet);
                }
            }

            xlWorkBook.SaveAs(ExcelPath, Excel.XlFileFormat.xlOpenXMLWorkbook, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlNoChange, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
            //add some text 
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
       

            WriteLog("Saved Excel Work book");
            KillWordAndExcelProcesses();
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

        private  void Generate_Word_Report(string SaveFolderName, string Xlsx_file_name, DataTable excelcalculation, List<string> Top5Maker, string DocxFileName)
        {
            // get image chart temp folder for each rv type
            string Imgsavefolder = img_temp + @"\" + SaveFolderName;

            // check if output folder is existing or not and create if not
            if (!Directory.Exists(Word_output))
                Directory.CreateDirectory(Word_output);

            if (!Directory.Exists(Word_output + @"\" + Xlsx_file_name))
                Directory.CreateDirectory(Word_output + @"\" + Xlsx_file_name);

            // generate docx file
            string docsaveLocation = Word_output + @"\" + Xlsx_file_name + @"\" + DocxFileName;
            WORD.Application wordApp = new WORD.Application();

            Document document = wordApp.Documents.Open(Word_report_template_file);


            // insert logo

            
            if (list_image_url.Count > 0)
            {
                Insert_Image_Chart("<INSERT_LOGO>", document, wordApp, list_image_url);

            }
            WriteLog("Inserted Logo");

            // add top 5 maker & model

            Replace_string_with_list(Top5Maker, "<TOP5_RV_MAKER>", document, wordApp, "\r\n");

            WriteLog("add top 5 maker & model");




            #region insert chart
            // insert map


            string TargetSave_CityMap = Imgsavefolder + url_cityMap;
            // 
            if(File.Exists(TargetSave_CityMap)) // case map is existing
            {
                Insert_Image_Chart("<CITY_MAP>", document, wordApp, TargetSave_CityMap);
                WriteLog("Inserted City Map");
            }
            else // if can't find map then just remove the target text
            {
                WriteLog("There is no map");
                Replace_string("<CITY_MAP>", "", wordApp);
            }

            // insert total Rv type Count Chart
            string TargetSave_url_numberofRVs = Imgsavefolder + url_numberofRVs;
            Insert_Image_Chart("<TOTAL_TYPE_CHART>", document, wordApp, TargetSave_url_numberofRVs);

            WriteLog("Inserted total Rv type Count Chart");

            // top 5 length
            string TargetSave_url_lengthof5RV = Imgsavefolder + url_lengthof5RV;
            Insert_Image_Chart("<TOP5_LENGTH_BAR_CHART>", document, wordApp, TargetSave_url_lengthof5RV);
            WriteLog("Inserted top 5 length Chart");

            // Average Daily Price/Night all
            string TargetSave_url_AvgDailyPriceNight = Imgsavefolder + url_AvgDailyPriceNight;
            Insert_Image_Chart("<AVG_DAILYPRICE_NIGHT_5_25_ALL_CHART>", document, wordApp, TargetSave_url_AvgDailyPriceNight);
            WriteLog("Inserted Average Daily Price/Night all Chart");

            // insert Price/night top 25 Chart
            string TargetSave_url_PriceNight = Imgsavefolder + url_PriceNight;
            Insert_Image_Chart("<TOP5_PRICE_NIGHT_SCRATER_CHART>", document, wordApp, TargetSave_url_PriceNight);
            WriteLog("Inserted Price/night top 5 Chart");

            // insert attributes chart
            string TargetSave_url_AttributesChart = Imgsavefolder + url_AcAhTable;
            Insert_Image_Chart("<TOP25_ATTRIBUTES_CHART>", document, wordApp, TargetSave_url_AttributesChart);
            WriteLog("Inserted attributes Chart");

            #endregion

            for (int i = 0; i < excelcalculation.Rows.Count; i++)
            {
                Replace_string(excelcalculation.Rows[i].Field<string>("FindString"), excelcalculation.Rows[i].Field<string>("ReplaceString"), wordApp);
            }
            WriteLog("Find string and replace");


            document.SaveAs2(docsaveLocation, ref misValue, ref misValue, ref misValue, ref misValue, ref misValue, ref misValue, ref misValue, ref misValue
                , ref misValue, ref misValue, ref misValue, ref misValue, ref misValue, ref misValue, ref misValue);

            // Close the document
            document.Close();
            // Quit the Word application
            wordApp.Quit();
            // clear image for next run

            WriteLog("Docx saved at :" + docsaveLocation);
            string[] Img_file = Directory.GetDirectories(img_temp);
            foreach(string file in Img_file)
            {
               Directory.Delete(file, true);
            }
   

            WriteLog("Clean template Image");

        }

        private static void Replace_string_with_list(List<string> ListStringToReplace, string string_to_find, Document document, WORD.Application WordApp, string lineBreak)
        {
            Document doc = document;


            doc.Activate();

            foreach (WORD.Range range in doc.StoryRanges)
            {

                string joinedText = String.Join(lineBreak, ListStringToReplace.ToArray());

                WORD.Find find = range.Find;
                object findText = string_to_find;

                object replacText = joinedText;
                object replace = WORD.WdReplace.wdReplaceAll;
                //object findWrap = WORD.WdFindWrap.wdFindContinue;

                if (replacText.ToString().Length > 254)
                {
                    WordApp.Application.Selection.Find.Execute(findText, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue,
                               misValue, misValue, misValue);
                    WordApp.Application.Selection.Text = (String)(replacText);
                    WordApp.Application.Selection.Collapse();
                }
                else
                {
                    find.Execute(ref findText, ref misValue, ref misValue, ref misValue, ref oFalse, ref misValue,
                        ref misValue, ref misValue, ref misValue, ref replacText,
                        ref replace, ref misValue, ref misValue, ref misValue, ref misValue);
                }


            }


        }

        private static void Insert_Image_Chart(string TexttoFind, Document Doc, WORD.Application WordApp, List<string> pictureLocation)
        {
            WORD.Range range = Doc.Content;
            range.Find.ClearFormatting();
            range.Find.Text = TexttoFind;
            range.Find.MatchCase = false;
            range.Find.MatchWholeWord = true;
            bool found = range.Find.Execute();

            if (found)
            {
                WORD.Range foundRange = range.Find.Parent;

                // Get the row below the found range

                int height = 30;
                foreach (string item in pictureLocation)
                {

                    System.Drawing.Image img = System.Drawing.Image.FromFile(item);

                    Double ratio = img.Width / img.Height;

                    InlineShape inlineShape = Doc.InlineShapes.AddPicture(item, ref misValue, ref misValue, foundRange);

                    inlineShape.Height = height;

                    inlineShape.Width = (int)(height * ratio);
                }

                // clear text

                Replace_string(TexttoFind, "", WordApp);
            }

        }


        private static void Insert_Image_Chart(string TexttoFind, Document Doc, WORD.Application WordApp, string pictureLocation)
        {
            WORD.Range range = Doc.Content;
            range.Find.ClearFormatting();
            range.Find.Text = TexttoFind;
            range.Find.MatchCase = false;
            range.Find.MatchWholeWord = true;
            bool found = range.Find.Execute();

            if (found)
            {
                WORD.Range foundRange = range.Find.Parent;

                // Get the row below the found range

                Doc.InlineShapes.AddPicture(pictureLocation, ref misValue, ref misValue, foundRange);
                // clear text

                Replace_string(TexttoFind, "", WordApp);
            }

        }

        private static void Replace_string(string string_to_find, string string_to_repace, WORD.Application WordApp)
        {
            Find find = WordApp.Selection.Find;
            find.Text = string_to_find;
            find.Replacement.Text = string_to_repace;
            find.Forward = true;
            find.Wrap = WdFindWrap.wdFindContinue;
            find.Format = false;
            find.MatchCase = false;
            find.MatchWholeWord = false;
            find.MatchWildcards = false;
            find.MatchSoundsLike = false;
            find.MatchAllWordForms = false;

            // Perform the Find and Replace
            object replace = WdReplace.wdReplaceAll;
            object missing = Type.Missing;
            find.Execute(FindText: missing, MatchCase: false, MatchWholeWord: true,
                         MatchWildcards: false, MatchSoundsLike: missing,
                         MatchAllWordForms: false, Forward: true,
                         Wrap: WdFindWrap.wdFindContinue, Format: false,
                         ReplaceWith: missing, Replace: replace);
        }

        private static void releaseObject(object obj)
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

        #endregion

        private void button1_Click_1(object sender, EventArgs e)
        {

        }

        private void button2_Click_1(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void btn_clearLog_Click(object sender, EventArgs e)
        {
            txt_log.Clear();
        }

        private void btn_openOutPut_Click(object sender, EventArgs e)
        {
            Process.Start("explorer.exe", txt_DocxOutPutPath.Text);
        }

        private void ChildForm_DataTableReturned(object sender, DataTableEventArgs e)
        {
            DataTable returnedDataTable = e.DataTable;

            //should put the power automate here
                this.Invoke(new MethodInvoker(delegate ()
                {
                    gmap.MaxZoom = 18;
                    gmap.MinZoom = 2;
                    gmap.Zoom = 8;
                    gmap.MapProvider = GMap.NET.MapProviders.GoogleMapProvider.Instance;
                    GMap.NET.GMaps.Instance.Mode = GMap.NET.AccessMode.ServerOnly;

                    //DataTable LatLongData = new DataTable();
                    //LatLongData.Columns.Add("LAT", typeof(double));
                    //LatLongData.Columns.Add("LONG", typeof(double));
                    //LatLongData.Rows.Add(33.588, -112.152);

                    // avg lat long
                    double AvgLat = returnedDataTable.AsEnumerable().Average(row => row.Field<double>("LAT"));
                    double AvgLong = returnedDataTable.AsEnumerable().Average(row => row.Field<double>("LONG"));


                    int numberOfPoints = 100;

                    // convert from mile to meter
                    double radiusInMeters = Double.Parse(txt_circleDiameter.Text) * 1609.344;

                    List<GMap.NET.PointLatLng> circlePoints = new List<GMap.NET.PointLatLng>();
                    double angle = 2 * Math.PI / numberOfPoints;
                    for (int i = 0; i < numberOfPoints; i++)
                    {
                        double lat = AvgLat + radiusInMeters / 111320d * Math.Sin(i * angle);
                        double lng = AvgLong + radiusInMeters / (111320d * Math.Cos(AvgLat * Math.PI / 180)) * Math.Cos(i * angle);
                        circlePoints.Add(new GMap.NET.PointLatLng(lat, lng));
                    }

                    GMap.NET.WindowsForms.GMapOverlay Overlay = new GMap.NET.WindowsForms.GMapOverlay("Overlay");
                    GMap.NET.WindowsForms.GMapPolygon circle = new GMap.NET.WindowsForms.GMapPolygon(circlePoints, "circle");

                    circle.Fill = new SolidBrush(Color.FromArgb(30, Color.Blue));  // Fill color
                    circle.Stroke = new Pen(Color.Blue, 1);

                    Overlay.Polygons.Add(circle);

                    gmap.Overlays.Add(Overlay);

                    // Create a GMapMarker for the center

                    // Add the marker to the overlay              


                    // add all location to overlay
                    foreach (DataRow dtrow in returnedDataTable.Rows)
                    {
                        GMap.NET.WindowsForms.GMapOverlay markers = new GMap.NET.WindowsForms.GMapOverlay("markers");
                        //GMap.NET.WindowsForms.GMapMarker marker = new GMap.NET.WindowsForms.Markers.GMarkerGoogle(
                        //                                                new GMap.NET.PointLatLng(dtrow.Field<double>("lat"), dtrow.Field<double>("long")),
                        //                                                            GMap.NET.WindowsForms.Markers.GMarkerGoogleType.red_small);
                        //Overlay = new GMap.NET.WindowsForms.GMapOverlay("markers");
                        GMap.NET.WindowsForms.GMapMarker marker = new GMap.NET.WindowsForms.Markers.GMarkerGoogle(
                                                                        new GMap.NET.PointLatLng(dtrow.Field<double>("lat"), dtrow.Field<double>("long")),
                                                                                    GMap.NET.WindowsForms.Markers.GMarkerGoogleType.red_small);
                        markers.Markers.Add(marker);
                        gmap.Overlays.Add(markers);

                    }


                    gmap.Position = new GMap.NET.PointLatLng(AvgLat, AvgLong);
                }
            ));


        }

        private void button1_Click(object sender, EventArgs e)
        {
            DataTable LatLongData = new DataTable();
            LatLongData.Columns.Add("LAT", typeof(double));
            LatLongData.Columns.Add("LONG", typeof(double));

            InputLatLong latLongForm = new InputLatLong(LatLongData);

            latLongForm.TopMost = true;
            latLongForm.DataTableReturned += ChildForm_DataTableReturned;
            latLongForm.ShowDialog();


           


        }

        private void LatLongForm_DataTableReturned(object sender, DataTableEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            KillWordAndExcelProcesses();
        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                if(TokenSource !=null)
                TokenSource.Cancel();
            }
            catch(Exception ex)
            {

            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            int OverLayCount = gmap.Overlays.Count;
            if (OverLayCount > 0)
            {             
                foreach( GMapOverlay overlay in gmap.Overlays.ToList() )
                {
                    gmap.Overlays.Remove(overlay);
                    gmap.Refresh();
                }    
              

            }

           

        }
    }
}
