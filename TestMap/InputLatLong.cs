using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RvAutoReport
{
    public partial class InputLatLong : Form
    {
        public DataTable latlongtable;

        public event EventHandler<DataTableEventArgs> DataTableReturned;
        public InputLatLong(DataTable latlong)
        {
            InitializeComponent();
            latlongtable = latlong;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void lv_latlong_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.V)
            {
                string clipboard = Clipboard.GetText().Trim();
                List<string> templist = clipboard.Replace("\r","").Split('\n').ToList();
               
                LoadDataTablefromListstring(templist);
            }
        }

        private void LoadDataTablefromListstring(List<string> ListLatLong)
        {
            latlongtable.Clear();
            foreach (string str in ListLatLong)
            {
                latlongtable.Rows.Add(str.Split('\t').ToList()[0] , str.Split('\t').ToList()[1] ) ;
            }
           
            lv_latlong.Items.Clear();
            if (latlongtable.Rows.Count > 0)
            {
               for(int i = 0;  i < latlongtable.Rows.Count; i++)
                {
                    ListViewItem listViewItem = lv_latlong.Items.Add(latlongtable.Rows[i].Field<double>("LAT").ToString() );
                    listViewItem.SubItems.Add(latlongtable.Rows[i].Field<double>("LONG").ToString());

                }
            }

            // distinct data

            latlongtable = latlongtable.AsEnumerable().GroupBy(dtrow => new
            {
                LAT = dtrow.Field<double>("LAT"),
                LONG = dtrow.Field<double>("LONG")
            }).Select(Group => Group.First()).CopyToDataTable();


            lbCount.Text = "Total :" + latlongtable.Rows.Count.ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OnDataTableReturned(latlongtable);
            this.Close();
        }

        protected virtual void OnDataTableReturned(DataTable dataTable)
        {
            DataTableReturned?.Invoke(this, new DataTableEventArgs(dataTable));
        }
    }

    public class DataTableEventArgs : EventArgs
    {
        public DataTable DataTable { get; }

        public DataTableEventArgs(DataTable dataTable)
        {
            DataTable = dataTable;
        }
    }
}
