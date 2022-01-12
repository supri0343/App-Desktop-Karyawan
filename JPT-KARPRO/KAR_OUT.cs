using MetroFramework.Forms;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Data;
using System.Data.Odbc;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Syncfusion.XlsIO;
using Spire.Xls;

namespace JPT_KARPRO
{
    public partial class KAR_OUT : Form
    {
        private DataTable dt = new DataTable();
        private DataTable deptdt = new DataTable();
        private SqlCommand cmd1;
        private SqlCommand cmd2;
        private SqlDataAdapter da;
        private DataSet ds;
        private SqlDataReader rd;
        private SqlDataAdapter da2;
        private DataSet ds2;
        private SqlDataReader rd2;
        private DataTable dt2 = new DataTable();


        public KAR_OUT()
        {
            InitializeComponent();
        }

        //akses database pc1
        const string connString = "Data Source=ANNNIIXX;Initial Catalog=JAPER-DATA;integrated security = true";
        SqlConnection koneksi = new SqlConnection("Data Source=ANNNIIXX;Initial Catalog=JAPER-DATA;integrated security = true");


        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

            koneksi.Open();
            string mulai = dateTimePicker1.Value.ToString(string.Format("yyyy-MM-dd"));
            string sampai = dateTimePicker2.Value.ToString(string.Format("yyyy-MM-dd"));
            cmd1 = new SqlCommand();
            cmd1 = koneksi.CreateCommand();
            cmd1.CommandText = "SELECT BIO.NPK,BIO.NAMA,BIO.BAGIAN,BIO.TGLLAHIR,BIO.TMK,BIO.TKK,BIO.KETERANGAN FROM [BIO] where (TKK BETWEEN ' " + mulai +" ' AND ' " + sampai +" ') ORDER by TKK asc";
            //cmd = new OdbcCommand(sql, connString);
            ds = new DataSet();
            da = new SqlDataAdapter(cmd1);
            da.Fill(ds, "data");
            //export datatable to excel
            DataTable t = ds.Tables[0];
            Workbook book = new Workbook();
            Worksheet sheet = book.Worksheets[0];
            sheet.InsertDataTable(t, true, 1, 1);
            sheet.Name = "Out";
            //auto collumn
            sheet.AllocatedRange.AutoFitColumns();
            sheet.AllocatedRange.AutoFitRows();
            //style
            sheet.Range["A1:G1"].Style.Color = Color.Gray;
            sheet.Range["A1:G1"].Style.Font.IsBold = true;
            sheet.FreezePanes(2, 1);

            System.Windows.Forms.SaveFileDialog saveDlg = new System.Windows.Forms.SaveFileDialog();
            saveDlg.InitialDirectory = @"C:\Users\Acer\Music";
            saveDlg.Filter = "Excel 97-2003 Workbook(*.xls)|*.xls|Excel Workbook(*.xlsx)|*xlsx";
            saveDlg.DefaultExt = ".xlsx";
            saveDlg.FilterIndex = 0;
            saveDlg.RestoreDirectory = true;
            saveDlg.FileName = "Laporan Out ";
            if (saveDlg.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveDlg.CheckPathExists)
            {

                //MessageBox.Show("Berhasil di Export");
                MessageBox.Show("Berhasil di Export");
                string path = saveDlg.FileName;
                book.SaveToFile(path);
                //book.Saved = true;
                if (MessageBox.Show("Apa anda ingin membuka file ?", "Export to Excel", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    System.Diagnostics.Process proc = new System.Diagnostics.Process();
                    proc.StartInfo.FileName = saveDlg.FileName;
                    proc.Start();

                }
            }
            koneksi.Close();
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            //declare variable from datetimepicker
            DateTime start = Convert.ToDateTime(dateTimePicker1.Value);
            DateTime end = Convert.ToDateTime(dateTimePicker2.Value);

            if (end >= start)
            {

                //use timespan to subtrack(mengurangi) waktu
                TimeSpan selisih = end.Subtract(start);
                //selisih tahun
                int totalyear = Convert.ToInt32(selisih.Days / 365);
                //selisih bulan
                int totalmonth = Convert.ToInt32(selisih.Days % 365 / 30);
                //selisih hari
                int totaldays = selisih.Days - totalyear * 365 - totalmonth * 30;

            }

            else
            {
                //jika tanggal salah maka isi data dibawah ini 
                dateTimePicker2.Value = DateTime.Now;
                MessageBox.Show("Harap Atur Tanggal Dengan Benar.", "Informasi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}
