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
using Microsoft.VisualBasic.FileIO;
using System.Data.Entity.Core.Common.CommandTrees.ExpressionBuilder;
using Button = System.Windows.Forms.Button;

namespace RvAutoReport
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }
        // from old script
        static string csv_input;
        static string xlsx_output;
        static string Logo_path;
        static string Word_output;
        static List<string> Word_report_template_file;

        static readonly string Report_Template_folder = Environment.CurrentDirectory + @"\template\";
        static readonly string img_temp = Environment.CurrentDirectory + @"\temp_image";
        static object misValue = System.Reflection.Missing.Value;
        static object oFalse = false;
        static object oTrue = true;
        static string Report_Type;

        // image template name , do not touch
        static string url_PriceNight = @"\priceNight.png";
        static string url_numberofRVs = @"\numberofRVs.png";
        static string url_lengthof5RV = @"\lengthof5RV.png";
        static string url_AvgDailyPriceNight = @"\AvgDailyPriceNight.png";
        static string url_AcAhTable = @"\AcAH.png";
        static string url_cityMap = @"\CityMap.png";

        static List<string> list_image_url = new List<string>();
        public System.Threading.CancellationTokenSource TokenSource;
        public DateTime Start_time;

        //  for new version // due to bug at export map image so not yet using these
        //public string ExcelPath = @"";
        public DataTable Excel_data;
        //public List<string> list_image_url = new List<string>();
        DataTable DataForWordReport = new DataTable() { TableName = "DataForWOrdReport" };
        public readonly string DataXmlPath = "Data.xml";
        public readonly string ConfigXmlPath = "Config.xml";

        private void Form1_Load(object sender, EventArgs e)
        {

            Word_report_template_file = new List<string>();

            string[] files = GetFileNamesFromFolder(Report_Template_folder, "*.docx");

            if (files.Length > 0)
            {
                cbb_SelectReport.DataSource = files;
                cbb_SelectReport.SelectedText = files.FirstOrDefault();
            }
            try
            {
                // load saved config data from config xml file if  file is existing
                WriteLog("Loading Config file");
                if (File.Exists(ConfigXmlPath))
                {
                    DataSet dtset = new DataSet();
                    dtset.ReadXml(ConfigXmlPath);

                    DataTable LoadFromXml = dtset.Tables[0];

                    if (LoadFromXml.Rows.Count > 0)
                    {

                        for (int i = 0; i < LoadFromXml.Rows.Count; i++)
                        {
                            string Config = LoadFromXml.Rows[i].Field<string>("CONFIG");
                            switch (Config)
                            {
                                case "CSV_PATH":

                                    csv_input = LoadFromXml.Rows[i].Field<string>("VALUE");
                                    txt_csvInput.Text = csv_input;
                                    break;

                                case "XLSX_PATH":
                                    xlsx_output = LoadFromXml.Rows[i].Field<string>("VALUE");
                                    txt_xlsxOutput.Text = xlsx_output;
                                    break;

                                case "DOCX_OUTPUT":
                                    Word_output = LoadFromXml.Rows[i].Field<string>("VALUE");
                                    txt_docxOutPut.Text = Word_output;
                                    break;

                                case "LOGO_PATH":
                                    Logo_path = LoadFromXml.Rows[i].Field<string>("VALUE");
                                    txt_logopath.Text = Logo_path;
                                    break;

                                case "ISMULTIREPORT":
                                    if (bool.Parse(LoadFromXml.Rows[i].Field<string>("VALUE")))
                                    {

                                        rd_runAllReport.Checked = true;
                                    }
                                    else
                                    {
                                        rd_runOneReport.Checked = true;
                                    }

                                    break;
                            }

                        }

                    }
                }

                else
                {
                    WriteLog("File Setting not found, Please change the settings manualy");
                }
            }
            catch (Exception ex)
            {
                WriteLog("Error Load Config, Please check the setting , re-config, save the retart the app");
                WriteLog(ex.ToString());
            }


        }

        private string[] GetFileNamesFromFolder(string folderPath, string pattern)
        {
            if (Directory.Exists(folderPath))
            {
                return Directory.GetFiles(folderPath, pattern)
                                .Select(Path.GetFileName).Where(file => !file.StartsWith("~"))
                                .ToArray();
            }
            else
            {
                Console.WriteLine("The specified folder does not exist.");
                return new string[0]; // Return an empty array if the folder does not exist
            }
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

        #region ignore
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
        #endregion

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
            if (!string.IsNullOrEmpty(txt_xlsxOutput.Text) && !string.IsNullOrEmpty(txt_csvInput.Text) && !string.IsNullOrEmpty(txt_docxOutPut.Text) && !string.IsNullOrEmpty(txt_logopath.Text))
            {
                try
                {
                    Word_report_template_file.Clear();
                    csv_input = txt_csvInput.Text;
                    xlsx_output = txt_xlsxOutput.Text;
                    Word_output = txt_docxOutPut.Text;
                    Logo_path = txt_logopath.Text;

                    // save the config file
                    File.Delete(ConfigXmlPath);
                    DataTable Config = new DataTable();
                    Config.TableName = "Config";
                    Config.Columns.Add("CONFIG");
                    Config.Columns.Add("VALUE");
                    Config.Rows.Add("XLSX_PATH", txt_xlsxOutput.Text);
                    Config.Rows.Add("CSV_PATH", txt_csvInput.Text);
                    Config.Rows.Add("DOCX_OUTPUT", txt_docxOutPut.Text);
                    Config.Rows.Add("MAP_CIRCLE_DIA", txt_circleDiameter.Text);
                    Config.Rows.Add("LOGO_PATH", txt_logopath.Text);

                    if (rd_runAllReport.Checked)
                    {
                        Config.Rows.Add("ISMULTIREPORT", true);

                    }
                    else
                    {
                        Config.Rows.Add("ISMULTIREPORT", false);

                    }
                    Config.WriteXml(ConfigXmlPath);
                    MessageBox.Show("Success !!!");
                }
                catch
                {
                    MessageBox.Show("Unknow error");
                }

            }
            else
            {
                MessageBox.Show("Please Input All Fields");
            }



            // reload variable



        }

        #region Copy From preview Script
        private void button4_Click(object sender, EventArgs e)
        {
            txt_log.Clear();
            ChangeCotrolStatus(false);
            WriteLog("Kill Excel , Word instance");
            KillWordAndExcelProcesses();

            foreach (string item in Word_report_template_file)
            {
                if (!File.Exists(Report_Template_folder + item))
                {
                    WriteLog("Word Template not found! Please check again ");
                    return;
                }
            }




            string[] files = Directory.GetFiles(xlsx_output, "*.xlsx");
            if (files.Length > 0)
            {
                WriteLog("Start Loop All file xlsx in folder");

                TokenSource = new System.Threading.CancellationTokenSource();
                CancellationToken token = TokenSource.Token;
                Task.Factory.StartNew(() =>

                {
                    try
                    {
                        Start_time = DateTime.Now;
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
                        }
                        WriteLog("Elapsed Time : " + Math.Round((DateTime.Now - Start_time).TotalSeconds, 0));
                        WriteLog("DONE!!!!!!!!!");
                        this.Invoke(new MethodInvoker(delegate
                        {
                            ChangeCotrolStatus(true);
                        }
    ));
                    }
                    catch (OperationCanceledException) { }
                    catch (ObjectDisposedException) { }
                    catch (Exception ex)
                    {
                        WriteLog(DateTime.Now + " " + ex.ToString());
                        this.Invoke(new MethodInvoker(delegate
                        {
                            ChangeCotrolStatus(true);
                        }
                        ));
                        //TokenSource.Cancel();
                    }

                }, token, TaskCreationOptions.LongRunning, TaskScheduler.Default);
            }
            else
            {
                WriteLog("No File !!");
            }



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
                        GMap.NET.WindowsForms.GMapMarker marker = new GMap.NET.WindowsForms.Markers.GMarkerGoogle(
                                                                        new GMap.NET.PointLatLng(dtrow.Field<double>("lat"), dtrow.Field<double>("long")),
                                                                                    GMap.NET.WindowsForms.Markers.GMarkerGoogleType.red_small);
                        markers.Markers.Add(marker);
                        gmap.Overlays.Add(markers);

                    }
                    gmap.Position = new GMap.NET.PointLatLng(AvgLat, AvgLong);
                    gmap.Refresh();
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

        private void ReadExcel(string ExcelPath)
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
            FullData = FullData.AsEnumerable().OrderByDescending(row => row.Field<double>("Current YTD Utilization")).CopyToDataTable();


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

                    // find city name
                    List<string> cityNametemp = xlWorkBook.Name.Split('-').ToList();
                    string CityName = cityNametemp[1];

                    // set active working sheet
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(shittt.Index);

                    // clear list logo image
                    list_image_url.Clear();

                    // get the RvType
                    string SheetName = shittt.Name.Replace("stat-", "");
                    string RvType = string.Empty;

                    // add class to Rv Type
                    if (SheetName == "A" || SheetName == "B" || SheetName == "C")
                    {
                        RvType = "Class " + SheetName;
                    }
                    else
                    {
                        RvType = SheetName;
                    }
                    WriteLog("RvType : " + RvType);

                    // create folder for template image file
                    string FolderPath = img_temp + @"\" + RvType;

                    if (!Directory.Exists(FolderPath))
                    {
                        Directory.CreateDirectory(FolderPath);
                    }
                    WriteLog("Image Folder Path : " + FolderPath);

                    // create Docx report file name  :  city name - RvType
                    string DocFileName = xlWorkBook.Name.Replace(CityName, CityName + "-" + RvType).Replace("xlsx", "docx");

                    // read accompanying data set and sort 
                    DataTable Data = ReadExcelFile(shittt.Name.Replace("stat-", "data-"), ExcelPath);

                    #region Map Making
                    // check if lat & long data is empty or not
                    DataTable RvData = new DataTable();
                    var checkifemptydata = Data.AsEnumerable()
                    .Where(row => !row.IsNull("LATITUDE") && !row.IsNull("LONGITUDE"));

                    if (checkifemptydata.Any())
                    {
                        RvData = checkifemptydata.CopyToDataTable();
                    }

                    // case lat long data is available, start creating map 
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

                            // clean map after exported image
                            foreach (GMapOverlay overlay in gmap.Overlays.ToList())
                            {
                                gmap.Overlays.Remove(overlay);
                                gmap.Refresh();
                            }
                            WriteLog("Map image exported");

                        }));

                    }
                    else
                    {
                        WriteLog("missing lat & long data");
                    }
                    #endregion

                    DataTable DataForWordReport = new DataTable("DataForWOrdReport");
                    DataForWordReport.Columns.Add("FindString");
                    DataForWordReport.Columns.Add("ReplaceString");

                    WriteLog(" get List top 5 Maker Name");

                    List<string> Top5MakerandModel = Data.AsEnumerable()
                                        .Select((row) => row.Field<string>("MAKE"))
                                        .Distinct().Take(5)
                                        .ToList();
                    for (int i = 0; i < Top5MakerandModel.Count; i++)
                    {
                        Top5MakerandModel[i] = (i + 1).ToString() + ". " + Top5MakerandModel[i];
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

                    //DataForWordReport.Rows.Add("<TOP5_AVG_UTI_IN_SEASON>", (xlWorkSheet.Cells[15, 2].Value * 100) + "%");

                    ////Average Utilization Off Season (Nov1 to Apr30) Top 5
                    //DataForWordReport.Rows.Add("<TOP5_AVG_UTI_OFF_SEASON>", (xlWorkSheet.Cells[16, 2].Value * 100) + "%");

                    ////Average Utilization In Season (May 1 to Oct 31) Top 25
                    //DataForWordReport.Rows.Add("<TOP25_AVG_UTI_IN_SEASON>", (xlWorkSheet.Cells[17, 2].Value * 100) + "%");

                    ////Average Utilization Off Season (Nov1 to Apr30) Top 25
                    //DataForWordReport.Rows.Add("<TOP25_AVG_UTI_OFF_SEASON>", (xlWorkSheet.Cells[18, 2].Value * 100) + "%");

                    //Average Utilization 2021 Top 5
                    string data = string.Empty;
                    data = xlWorkSheet.Cells[11, 2].Value.ToString();
                    DataForWordReport.Rows.Add("<TOP5_AVG_UTI_2021>", (double.Parse(xlWorkSheet.Cells[11, 2].Value.ToString()) * 100) + "%");

                    //Average Utilization 2022 Top 5
                    data = xlWorkSheet.Cells[12, 2].Value.ToString();
                    DataForWordReport.Rows.Add("<TOP5_AVG_UTI_2022>", (double.Parse(xlWorkSheet.Cells[12, 2].Value.ToString()) * 100) + "%");

                    //Average Utilization of top 5 YTD
                    data = xlWorkSheet.Cells[13, 2].Value.ToString();
                    DataForWordReport.Rows.Add("<TOP5_AVG_UTI_YTD>", (double.Parse(xlWorkSheet.Cells[13, 2].Value.ToString()) * 100) + "%");

                    //Average Utilization of Top 25 YTD
                    data = xlWorkSheet.Cells[14, 2].Value.ToString();
                    DataForWordReport.Rows.Add("<TOP25_AVG_UTI_YTD>", (double.Parse(xlWorkSheet.Cells[14, 2].Value.ToString()) * 100) + "%");

                    //Future utilization 30 Days
                    data = xlWorkSheet.Cells[19, 2].Value.ToString();
                    DataForWordReport.Rows.Add("<FUTURE_UTI_30>", (double.Parse(xlWorkSheet.Cells[19, 2].Value.ToString()) * 100) + "%");

                    //Future utilization 60 Days
                    data = xlWorkSheet.Cells[20, 2].Value.ToString();
                    DataForWordReport.Rows.Add("<FUTURE_UTI_60>", (double.Parse(xlWorkSheet.Cells[20, 2].Value.ToString()) * 100) + "%");

                    //Future utilization 90 Days
                    data = xlWorkSheet.Cells[21, 2].Value.ToString();
                    DataForWordReport.Rows.Add("<FUTURE_UTI_90>", (double.Parse(xlWorkSheet.Cells[21, 2].Value.ToString()) * 100) + "%");

                    //  Potential Annual
                    data = xlWorkSheet.Cells[22, 2].Value.ToString();
                    DataForWordReport.Rows.Add("<POTENTIAL_ANNUAL>", "$" + xlWorkSheet.Cells[22, 2].Value);

                    //  Average nightly price of top 5 RVs

                    // sort data again by UTIL 2022
                    Data = Data.AsEnumerable().OrderByDescending(row => row.Field<double>("UTIL 2022")).CopyToDataTable();
                    //Data.DefaultView.Sort = "[UTIL 2022] DESC";

                    DataTable Data2 = Data.DefaultView.ToTable();

                    double AvgTop5NightlyPrice = Math.Round(Data2.AsEnumerable().Take(5).Select(row => row.Field<double>("PRICE/NIGHT")).Average(), 2);

                    DataForWordReport.Rows.Add("<AVG_NP_TOP5>", AvgTop5NightlyPrice);

                    //  Average nightly price calculation for last table

                    DataForWordReport.Rows.Add("<AVG_NP_TOP5_025>", Math.Round(AvgTop5NightlyPrice * (0.25 * 365)));

                    DataForWordReport.Rows.Add("<AVG_NP_TOP5_026>", Math.Round(AvgTop5NightlyPrice * (0.26 * 365)));

                    DataForWordReport.Rows.Add("<AVG_NP_TOP5_050>", Math.Round(AvgTop5NightlyPrice * (0.50 * 365)));

                    DataForWordReport.Rows.Add("<AVG_NP_TOP5_051>", Math.Round(AvgTop5NightlyPrice * (0.51 * 365)));

                    DataForWordReport.Rows.Add("<AVG_NP_TOP5_075>", Math.Round(AvgTop5NightlyPrice * (0.75 * 365)));

                    DataForWordReport.Rows.Add("<AVG_NP_TOP5_076>", Math.Round(AvgTop5NightlyPrice * (0.76 * 365)));

                    DataForWordReport.Rows.Add("<AVG_NP_TOP5_100>", Math.Round(AvgTop5NightlyPrice * 365));



                    #endregion


                    DataForWordReport.WriteXml("data.xml");


                    WriteLog("Creating Chart");
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
                        , "RV Count", XlChartType.xl3DColumnClustered, 5, 650, 500, 300, FolderPath, url_numberofRVs, Office.MsoThemeColorIndex.msoThemeColorLight1, Color.Black, Excel.XlRgbColor.rgbBlack);

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

                    double avgtop5 = (double)xlWorkSheet.Cells[9, 2].Value;

                    double avgtop25 = (double)xlWorkSheet.Cells[10, 2].Value;

                    double avgAll = Data.AsEnumerable().Average(row => row.Field<double>("PRICE/NIGHT"));

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
                    DataTable AcAHTable = new DataTable();
                    AcAHTable.Columns.Add("TYPE");
                    AcAHTable.Columns.Add("VALUE");

                    DataTable dt_forAttributes = null;
                    var Checkdt_forAttribute = Data.AsEnumerable().Take(25);
                    if (Checkdt_forAttribute.Any())
                    {
                        dt_forAttributes = Checkdt_forAttribute.CopyToDataTable();
                    }
                    if (dt_forAttributes != null)
                    {
                        double Top25PETFRIENDELY = dt_forAttributes.AsEnumerable().Sum(row => row.Field<double>("PET FRIENDELY")) / dt_forAttributes.Rows.Count;
                        double Top25TAILGATEFRIENDELY = dt_forAttributes.AsEnumerable().Sum(row => row.Field<double>("TAILGATE FRIENDELY")) / dt_forAttributes.Rows.Count;
                        double Top25SMOKINGALLOWED = dt_forAttributes.AsEnumerable().Sum(row => row.Field<double>("SMOKING ALLOWED")) / dt_forAttributes.Rows.Count;
                        double Top25FESTIVALFRIENDLY = dt_forAttributes.AsEnumerable().Sum(row => row.Field<double>("FESTIVAL FRIENDLY")) / dt_forAttributes.Rows.Count;
                        double Top25GENERATOR = dt_forAttributes.AsEnumerable().Sum(row => row.Field<double>("GENERATOR")) / dt_forAttributes.Rows.Count;
                        AcAHTable.Rows.Add("PET FRIENDELY", Top25PETFRIENDELY);
                        AcAHTable.Rows.Add("TAILGATE FRIENDELY", Top25TAILGATEFRIENDELY);
                        AcAHTable.Rows.Add("SMOKING ALLOWED", Top25SMOKINGALLOWED);
                        AcAHTable.Rows.Add("FESTIVAL FRIENDLY", Top25FESTIVALFRIENDLY);
                        AcAHTable.Rows.Add("GENERATOR", Top25GENERATOR);
                    }
                    else
                    {
                        AcAHTable.Rows.Add("PET FRIENDELY", 0);
                        AcAHTable.Rows.Add("TAILGATE FRIENDELY", 0);
                        AcAHTable.Rows.Add("SMOKING ALLOWED", 0);
                        AcAHTable.Rows.Add("FESTIVAL FRIENDLY", 0);
                        AcAHTable.Rows.Add("GENERATOR", 0);
                    }



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

                    foreach (string rv in Top5MakerandModel)
                    {
                        string LogoLocation = Logo_path + @"\" + rv.Substring(2).Trim() + ".png";
                        if (File.Exists(LogoLocation))
                        {
                            list_image_url.Add(LogoLocation);
                        }
                    }

                    // insert logo to sheet
                    if (list_image_url.Count > 0)
                    {
                        int x = 500;
                        int y = 0;
                        int width = 0;
                        int height = 50;
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


                    WriteLog("Generating Word report ");
                    // generate docx report file

                    foreach (string reportFile in Word_report_template_file)
                    {
                        Report_Type = reportFile.Replace(".docx", "");
                        string TemplateFile = Report_Template_folder + reportFile;
                        Generate_Word_Report(RvType, xlWorkBook.Name.Replace(".xlsx", ""), DataForWordReport, Top5MakerandModel, DocFileName, TemplateFile);

                    }


                    string[] Img_file = Directory.GetDirectories(img_temp);
                    foreach (string file in Img_file)
                    {
                        Directory.Delete(file, true);
                    }


                    WriteLog("Cleaned template Image");

                    // release excel work sheet
                    releaseObject(xlWorkSheet);
                }
            }
            // save excel workbook
            xlWorkBook.SaveAs(ExcelPath, Excel.XlFileFormat.xlOpenXMLWorkbook, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlNoChange, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
            // release work app
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

        private void Generate_Word_Report(string SaveFolderName, string Xlsx_file_name, DataTable excelcalculation, List<string> Top5Maker, string DocxFileName, string WordFile)
        {
            // get image chart temp folder for each rv type
            string Imgsavefolder = img_temp + @"\" + SaveFolderName;

            // check if output folder is existing or not and create if not
            string OutputFolder = Word_output + @"\" + Report_Type + @"\" + Xlsx_file_name;
            if (!Directory.Exists(OutputFolder))
                Directory.CreateDirectory(OutputFolder);

            // generate docx file name
            string docsaveLocation = OutputFolder + @"\" + DocxFileName;
            WORD.Application wordApp = new WORD.Application();

            Document document = wordApp.Documents.Open(WordFile);


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
            if (File.Exists(TargetSave_CityMap)) // case map is existing
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
            Process.Start("explorer.exe", txt_docxOutPut.Text);
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
                if (TokenSource != null)
                    TokenSource.Cancel();
            }
            catch (Exception ex)
            {

            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            int OverLayCount = gmap.Overlays.Count;
            if (OverLayCount > 0)
            {
                foreach (GMapOverlay overlay in gmap.Overlays.ToList())
                {
                    gmap.Overlays.Remove(overlay);
                    gmap.Refresh();
                }


            }



        }

        private void ChangeCotrolStatus(bool status)
        {
            foreach (Control ctrl in tabPage1.Controls)
            {
                if (ctrl.GetType() == typeof(Button))
                {
                    ctrl.Enabled = status;
                }
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            txt_log.Clear();
            ChangeCotrolStatus(false);
            string[] files = Directory.GetFiles(csv_input, "*.csv");

            if (files.Length > 0)
            {
                WriteLog("Start Loop All file csv in folder");

                TokenSource = new System.Threading.CancellationTokenSource();
                CancellationToken token = TokenSource.Token;
                Task.Factory.StartNew(() =>

                {
                    //try
                    //{
                    Start_time = DateTime.Now;
                    foreach (string file in files)
                    {
                        if (!file.Contains("~$")) // ignore the excel temp file
                        {

                            CSVtoXlsx(file);

                            if (token.IsCancellationRequested)
                            {
                                token.ThrowIfCancellationRequested();
                            }
                        }
                    }
                    WriteLog("Elapsed Time : " + Math.Round((DateTime.Now - Start_time).TotalSeconds));
                    WriteLog("DONE!!!!!!!!!");
                    this.Invoke(new MethodInvoker(delegate
                    {
                        ChangeCotrolStatus(true);
                    }
                    ));

                    //}
                    //catch (OperationCanceledException) { }
                    //catch (ObjectDisposedException) { }
                    //catch (Exception ex)
                    //{
                    //    WriteLog(DateTime.Now + " " + ex.ToString());
                    //    this.Invoke(new MethodInvoker(delegate
                    //    {
                    //        ChangeCotrolStatus(true);
                    //    }
                    //    ));
                    //    //TokenSource.Cancel();
                    //}

                }, token, TaskCreationOptions.LongRunning, TaskScheduler.Default);
            }
            else
            {
                WriteLog("No file!");
            }


        }


        private void CSVtoXlsx(string csvFilePath)
        {
            KillWordAndExcelProcesses();

            // test replace power automate 

            WriteLog("load csv datatable");
            DataTable Data = LoadCSV(csvFilePath);

            // generate xlsx path
            string ExcelPath = xlsx_output + @"\" + csvFilePath.Substring(csvFilePath.LastIndexOf(@"\") + 1).Replace(".csv", ".xlsx");

            // ignore empty type
            Data = Data.AsEnumerable().Where(rn => !string.IsNullOrEmpty(rn.Field<string>("TYPE"))).CopyToDataTable();

            // clean speical character (%,$) so data can be convert to decimal
            WriteLog("convert data");
            foreach (DataRow dtrow in Data.Rows)
            {
                string CurrentYTDUtilization = dtrow["Current YTD Utilization"].ToString().Replace("%", "");
                string PRICENIGHT = dtrow["PRICE/NIGHT"].ToString().Replace("$", "");
                string UTIL2021 = dtrow["UTIL 2021"].ToString().Replace("%", "");
                string UTIL2022 = dtrow["UTIL 2022"].ToString().Replace("%", "");
                string UTILFuture30 = dtrow["UTIL Future 30"].ToString().Replace("%", "");
                string UTILFuture60 = dtrow["UTIL Future 60"].ToString().Replace("%", "");
                string UTILFuture90 = dtrow["UTIL Future 90"].ToString().Replace("%", "");

                dtrow["Current YTD Utilization"] = CurrentYTDUtilization;
                dtrow["PRICE/NIGHT"] = PRICENIGHT;
                dtrow["UTIL 2021"] = UTIL2021;
                dtrow["UTIL 2022"] = UTIL2022;
                dtrow["UTIL Future 30"] = UTILFuture30;
                dtrow["UTIL Future 60"] = UTILFuture60;
                dtrow["UTIL Future 90"] = UTILFuture90;

            }

            // create excel instance 
            WriteLog("create excel instance");
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            xlApp = new Excel.Application();
            xlApp.DisplayAlerts = false;


            // create the first sheet with full data
            xlWorkBook = xlApp.Workbooks.Add();
            xlWorkSheet = (Worksheet)xlWorkBook.Worksheets[1];
            // Set the name of the worksheet
            xlWorkSheet.Name = "OriginalData";
            WriteLog("create excel sheet : " + xlWorkSheet.Name);
            // write data to original sheet
            WriteLog("write data to  sheet" + xlWorkSheet.Name);
            WriteDataTableToExcelSheet(xlWorkSheet, Data, 1, 1);


            // get list EV type
            WriteLog("get Rv list");
            List<string> RvList = Data.AsEnumerable().Select(row => row.Field<string>("TYPE")).Distinct().ToList();

            RvList = RvList.OrderByDescending(s => s).ToList();

            if (RvList.Count > 0)
            {
                // loop thru each Rv founded then create 2 sheet rvName-data RvName-stat
                WriteLog("start loop tru Rv list");
                foreach (string rv in RvList)
                {
                    // get current sheet
                    Worksheet currentSheet = (Worksheet)xlApp.ActiveSheet;

                    // create  rvName-data sheet
                    Excel.Worksheet NewSheetData = (Worksheet)xlWorkBook.Worksheets.Add(Before: currentSheet);
                    NewSheetData.Name = "data-" + rv;
                    WriteLog("created sheet " + NewSheetData.Name);
                    // filter RV data
                    DataTable rvData = Data.AsEnumerable().Where(clm => clm.Field<string>("TYPE") == rv).CopyToDataTable();
                    // sort by Current YTD Ulti
                    rvData = rvData.AsEnumerable().OrderByDescending(row => decimal.Parse(row.Field<string>("Current YTD Utilization"))).CopyToDataTable();

                    //rvData.DefaultView.Sort = "[Current YTD Utilization] DESC";
                    //rvData = rvData.DefaultView.ToTable();

                    // write data to worksheet 
                    WriteLog("Write data to " + NewSheetData.Name);
                    WriteDataTableToExcelSheet(NewSheetData, rvData, 1, 1);

                    currentSheet = (Worksheet)xlApp.ActiveSheet;

                    // create stat worksheet
                    Excel.Worksheet NewSheetStat = (Worksheet)xlWorkBook.Worksheets.Add(Before: currentSheet);
                    NewSheetStat.Name = "stat-" + rv;
                    WriteLog("created sheet " + NewSheetStat.Name);

                    // calculate data for  -stat sheet
                    DataTable xlData = new DataTable();
                    xlData.Columns.Add("NAME");
                    xlData.Columns.Add("VALUE");

                    int TotalRVCount = Data.Rows.Count;
                    int TotalRVTypeCount = rvData.Rows.Count;

                    WriteLog("start calculate data ");
                    //Total RVs in the analysis area : count all row in "Data" table 
                    xlData.Rows.Add("Total RVs in the analysis area", TotalRVCount);

                    // Total Class ClassName As in the  analysis area : Count all row in Type Filterd table 
                    xlData.Rows.Add("Total Class " + rv + " As in the  analysis area", TotalRVTypeCount);

                    // Percent of  Total RVs in the analysis area that are ClassName : rvData / Data
                    xlData.Rows.Add("Percent of  Total RVs in the analysis area that are Class " + rv, Math.Round((double)TotalRVTypeCount / (double)TotalRVCount * 100, 0) + "%");

                    //Average Age of all Class ClassName As in the analysis area, count in filtered Rvtype with Year > 0
                    int AvgAgeOfallClass = (int)Math.Round(rvData.AsEnumerable().Select(row => int.Parse(row.Field<string>("YEAR"))).Where(value => value > 0).Average(), 0);
                    xlData.Rows.Add("Average Age of all Class " + rv + " As in the analysis area", AvgAgeOfallClass);

                    //Average Age of the Top 5 Class ClassName As in the analyzed area
                    int AvgAgeOfTop5 = (int)Math.Round(rvData.AsEnumerable().Take(5).Select(row => int.Parse(row.Field<string>("YEAR"))).Where(value => value > 0).Average(), 0);
                    xlData.Rows.Add("Average Age of the Top 5 Class " + rv + " As in the analyzed area", AvgAgeOfTop5);

                    // Average Age of the Top 25 Class ClassName As in the analysis area
                    int AvgAgeOfTop25 = (int)Math.Round(rvData.AsEnumerable().Take(25).Select(row => int.Parse(row.Field<string>("YEAR"))).Where(value => value > 0).Average(), 0);
                    xlData.Rows.Add("Average Age of the Top 25 Class " + rv + " As in the analyzed area", AvgAgeOfTop25);

                    //Average length of top 5 RVs
                    int AvglengthOfTop5 = (int)Math.Round(rvData.AsEnumerable().Take(5).Select(row => decimal.Parse(row.Field<string>("LENGTH"))).Where(value => value > 0).Average(), 0);
                    xlData.Rows.Add("Average length of top 5 RVs", AvglengthOfTop5);

                    //Average length of top 25 RVs
                    int AvglengthOfTop25 = (int)Math.Round(rvData.AsEnumerable().Take(25).Select(row => decimal.Parse(row.Field<string>("LENGTH"))).Where(value => value > 0).Average(), 0);
                    xlData.Rows.Add("Average length of top 25 RVs", AvglengthOfTop25);

                    //Average nightly price of top 5 RVs
                    decimal NightlyPriceTop5 = Math.Round(rvData.AsEnumerable().Take(5).Select(row => decimal.Parse(row.Field<string>("PRICE/NIGHT"))).Average(), 2);
                    xlData.Rows.Add("Average nightly price of top 5 RVs", "$" + NightlyPriceTop5);

                    //Average nightly price of top 25 RVs
                    decimal NightlyPriceTop25 = Math.Round(rvData.AsEnumerable().Take(25).Select(row => decimal.Parse(row.Field<string>("PRICE/NIGHT").Replace("$", ""))).Average(), 2);
                    xlData.Rows.Add("Average nightly price of top 25 RVs", "$" + NightlyPriceTop25);

                    //Average Utilization 2021 Top 5
                    // sort by UTIL 2021 desc
                    DataTable dtFor_AvgUtli2021 = rvData.AsEnumerable().OrderByDescending(row => decimal.Parse(row.Field<string>("UTIL 2021"))).CopyToDataTable(); ;
                    // average by top5
                    decimal AvgUtli2021 = Math.Round(dtFor_AvgUtli2021.AsEnumerable().Take(5).Select(row => decimal.Parse(row.Field<string>("UTIL 2021"))).Average(), 0);
                    xlData.Rows.Add("Average Utilization 2021 Top 5", AvgUtli2021 + "%");

                    //Average Utilization 2022 Top 5
                    // sort by UTIL 2022 desc
                    DataTable dtFor_AvgUtli2022 = rvData.AsEnumerable().OrderByDescending(row => decimal.Parse(row.Field<string>("UTIL 2022"))).CopyToDataTable();
                    // average by top5
                    decimal AvgUtli2022 = Math.Round(dtFor_AvgUtli2022.AsEnumerable().Take(5).Select(row => decimal.Parse(row.Field<string>("UTIL 2022"))).Average(), 0);
                    xlData.Rows.Add("Average Utilization 2022 Top 5", AvgUtli2022 + "%");

                    //Average Utilization of top 5 YTD
                    decimal AvgUtliYTDTop5 = Math.Round(rvData.AsEnumerable().Take(5).Select(row => decimal.Parse(row.Field<string>("Current YTD Utilization"))).Average(), 0);
                    xlData.Rows.Add("Average Utilization of top 5 YTD", AvgUtliYTDTop5 + "%");
                    //Average Utilization of top 25 YTD
                    decimal AvgUtliYTDTop25 = Math.Round(rvData.AsEnumerable().Take(25).Select(row => decimal.Parse(row.Field<string>("Current YTD Utilization"))).Average(), 0);
                    xlData.Rows.Add("Average Utilization of top 25 YTD", AvgUtliYTDTop25 + "%");


                    // this calculation is not using, just leave it here to not break the second phase

                    xlData.Rows.Add("Average Utilization In Season (May 1 to Oct 31) Top 5", "%");
                    xlData.Rows.Add("Average Utilization Off Season (Nov1 to Apr30) Top 5", "%");
                    xlData.Rows.Add("Average Utilization In Season (May 1 to Oct 31) Top 25", "%");
                    xlData.Rows.Add("Average Utilization Off Season (Nov1 to Apr30) Top 25", "%");

                    //Future utilization 30 Days
                    decimal AvgFutu30UtliTop5 = Math.Round(rvData.AsEnumerable().Take(5).Select(row => decimal.Parse(row.Field<string>("UTIL Future 30").Replace("%", ""))).Average(), 0);
                    xlData.Rows.Add("Future utilization 30 Days", AvgFutu30UtliTop5 + "%");

                    //Future utilization 60 Days
                    decimal AvgFutu60UtliTop5 = Math.Round(rvData.AsEnumerable().Take(5).Select(row => decimal.Parse(row.Field<string>("UTIL Future 60").Replace("%", ""))).Average(), 0);
                    xlData.Rows.Add("Future utilization 60 Days", AvgFutu60UtliTop5 + "%");

                    //Future utilization 90 Days
                    decimal AvgFutu90UtliTop5 = Math.Round(rvData.AsEnumerable().Take(5).Select(row => decimal.Parse(row.Field<string>("UTIL Future 90").Replace("%", ""))).Average(), 0);
                    xlData.Rows.Add("Future utilization 90 Days", AvgFutu90UtliTop5 + "%");

                    //Potential Annual Revenue for the top 5 based on 2022 Utilization
                    // sort by UTIL 2022 desc

                    //DataTable dtFor_Potential = rvData.AsEnumerable().OrderByDescending(row => decimal.Parse(row.Field<string>("UTIL 2022")))
                    //    .Where(Row => Row.Field<string>("UTIL 2022") != "0").Take(5).CopyToDataTable();

                    DataTable dtFor_Potential = null;
                    var Check = rvData.AsEnumerable().OrderByDescending(row => decimal.Parse(row.Field<string>("UTIL 2022"))).Where(Row => Row.Field<string>("UTIL 2022") != "0");

                    if (Check.Any())
                    {
                        dtFor_Potential = Check.CopyToDataTable();
                    }
                    if (dtFor_Potential != null)
                    {
                        dtFor_Potential = dtFor_Potential.AsEnumerable().Take(5).CopyToDataTable();
                        // avg of Top5 Ulti2022
                        decimal avgTop5Ulti2022 = Math.Round(dtFor_Potential.AsEnumerable().Select(row => decimal.Parse(row.Field<string>("UTIL 2022").Replace("%", ""))).Average(), 0);
                        // avg of Top5 Price/Night
                        decimal avgTop5PriceNight = Math.Round(dtFor_Potential.AsEnumerable().Select(row => decimal.Parse(row.Field<string>("PRICE/NIGHT").Replace("$", ""))).Average(), 2);
                        decimal PAR = Math.Round((365 * avgTop5Ulti2022 / 100) * avgTop5PriceNight);
                        xlData.Rows.Add("Potential Annual Revenue for the top 5 based on 2022 Utilization", "$" + PAR);
                    }
                    xlData.Rows.Add("Potential Annual Revenue for the top 5 based on 2022 Utilization", "Not Enough Data");



                    WriteLog("start calculated data  to sheet " + NewSheetStat.Name);
                    // write data to stat worksheet
                    WriteDataTableToExcelSheet(NewSheetStat, xlData, 1, 1, false);

                    Excel.Range columnRange = NewSheetStat.Columns[1];
                    columnRange.EntireColumn.AutoFit();

                }
            }

            xlWorkBook.SaveAs(ExcelPath, Excel.XlFileFormat.xlOpenXMLWorkbook, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlNoChange, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
            WriteLog("xlsx file Created : " + ExcelPath);
            // release work app
            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }

        static void WriteDataTableToExcelSheet(Worksheet worksheet, DataTable dataTable, int startRow, int startColumn, bool isWriteHeader = true)
        {
            // Write the column headers to the worksheet
            if (isWriteHeader)
            {
                for (int col = 0; col < dataTable.Columns.Count; col++)
                {
                    worksheet.Cells[startRow, startColumn + col] = dataTable.Columns[col].ColumnName;
                }
                startRow = startRow + 1;
            }



            // Write the data rows to the worksheet
            for (int row = 0; row < dataTable.Rows.Count; row++)
            {
                for (int col = 0; col < dataTable.Columns.Count; col++)
                {
                    worksheet.Cells[startRow + row, startColumn + col] = dataTable.Rows[row][col];
                }
            }
        }


        private DataTable LoadCSV(string filePath)
        {
            DataTable dataTable = new DataTable();

            // Read the CSV file using TextFieldParser
            using (TextFieldParser parser = new TextFieldParser(filePath))
            {
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(","); // Set the delimiter used in your CSV file (e.g., comma)

                // Read the first line as header and create DataTable columns
                string[] headers = parser.ReadFields();
                foreach (string header in headers)
                {
                    dataTable.Columns.Add(header);
                }

                // Read and add each row to the DataTable
                while (!parser.EndOfData)
                {
                    string[] fields = parser.ReadFields();
                    dataTable.Rows.Add(fields);
                }
            }

            return dataTable;
        }

        private void rd_runAllReport_CheckedChanged(object sender, EventArgs e)
        {
            if (rd_runAllReport.Checked == true)
            {
                // load all docx file from template folder 
                cbb_SelectReport.Enabled = false;
                Word_report_template_file.Clear();
                string[] files = GetFileNamesFromFolder(Report_Template_folder, "*.docx");
                if (files.Length > 0)
                {
                    WriteLog("Mutlti report mode");
                    foreach (string DocFile in files)
                    {

                        if (!DocFile.StartsWith("~"))
                        {
                            Word_report_template_file.Add(DocFile);
                            WriteLog("Report File : " + DocFile);
                        }

                    }
                }
                else
                {
                    WriteLog("WARNING : Not found any Report !!! , please put atleast one report at the template Folder ");
                }


            }
            else
            {
                rd_runOneReport.Checked = true;



            }
        }

        private void rd_runOneReport_CheckedChanged(object sender, EventArgs e)
        {
            if (rd_runOneReport.Checked == true)
            {
                WriteLog("Single report mode");
                Word_report_template_file.Clear();

                Word_report_template_file.Add(cbb_SelectReport.SelectedItem.ToString());
                WriteLog("Report File : " + cbb_SelectReport.SelectedItem.ToString());
                cbb_SelectReport.Enabled = true;
            }



        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }
    }
}
