using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MetroFramework.Forms;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using Syncfusion.XlsIO;
using Spire.Xls;

namespace JPT_KARPRO
{
    public partial class OUT : Form
    {
        private DataTable dt = new DataTable();
        private DataTable deptdt = new DataTable();
        private SqlCommand cmd1;
        private SqlCommand cmd2;
        private SqlDataAdapter da;
        private DataSet ds;
        private SqlDataReader rd;


        public string npkterpilih;
        public OUT()
        {
       
            InitializeComponent();
            tampildata();
        }

        const string connString = "Data Source=ANNNIIXX;Initial Catalog=JAPER-DATA;integrated security = true";
        SqlConnection koneksi = new SqlConnection("Data Source=ANNNIIXX;Initial Catalog=JAPER-DATA;integrated security = true");

        private DataView myView;
        private DataTable dataTable = new DataTable();

        private void tampildata()
        {

            //SQL STATEMENT
            koneksi.Open();
            cmd1 = new SqlCommand();
            cmd1 = koneksi.CreateCommand();
            cmd1.CommandText = "SELECT TOP 20 BIO.NPK,BIO.NAMA,BIO.JK,BIO.TGLLAHIR,BIO.PDDK,BIO.AGAMA,BIO.TMK,BIO.USIA,BIO.TKK,BIO.ALAMAT,BIO.KABUPATEN,BIO.DOMISILI,BIO.KTP,BIO.IBU,BIO.HP,BIO.STATUS,BIO.KETERANGAN FROM [BIO] where tkk is not null order by tkk desc";
            //cmd = new OdbcCommand(sql, connString);

            ds = new DataSet();
            da = new SqlDataAdapter(cmd1);
            da.Fill(ds, "DATA kar");
            dataGridView1.DataSource = ds;
            dataGridView1.DataMember = "DATA kar";



            dataGridView1.Refresh();
            

            cmd1 = new SqlCommand();
            cmd1 = koneksi.CreateCommand();
            cmd1.CommandText = "SELECT COUNT(*) As count FROM [BIO] where TKK is NOT null";
            Int32 jum = (Int32)cmd1.ExecuteScalar() + 30 ;
            label9.Text = jum.ToString();
            //ds = new DataSet();
            //da = new SqlDataAdapter(cmd1);
            //da.Fill(ds);
            //label9.Text = ds.Tables[0].Rows[0]["count"].ToString();
            
            koneksi.Close();

            //int numRows = dataGridView1.Rows.Count + 29;
            //label9.Text = numRows.ToString();

        }
        private void update()
        {
            SqlConnection conn = new SqlConnection(connString);
            string sql = "UPDATE [BIO] SET nama='" + NAMA1.Text + "',jk='" + JK1.Text + "',tgllahir='" + LAHIR1.Value.ToString(string.Format("yyyy-MM-dd")) + "',pddk='" + PDD1.Text + "',agama='" + AGAMA.Text + "',tmk='" + TMK1.Value.ToString(string.Format("yyyy-MM-dd")) + "',usia='" + USIA1.Text + "',tkk='" + tkk.Value.ToString(string.Format("yyyy-MM-dd")) + "',alamat='" + ALAMAT1.Text + "',kabupaten='" + KAB1.Text + "',domisili='" + DOMI.Text + "',ktp='" + KTP1.Text + "',ibu='" + IBU1.Text + "',hp='" + hp.Text + "',status='" + sta.Text + "',KETERANGAN='" + KET.Text + "' where npk='" + NPK1.Text + "'";
            SqlCommand cmd = new SqlCommand(sql, conn);
            try
            {
                conn.Open();
                SqlDataAdapter da = new SqlDataAdapter(cmd)
                {
                    UpdateCommand = conn.CreateCommand()
                };
                da.UpdateCommand.CommandText = sql;
                if (da.UpdateCommand.ExecuteNonQuery() > 0)
                {

                    MessageBox.Show(@"Successfully Updated");
                }
                conn.Close();

                //REFRESH DATA
                DataTable dt = new DataTable();
                dt.Rows.Clear();
                
                reset();
              

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                koneksi.Close();
            }
        }

        private void reset()
        {
            NPK1.Text = "";
            NAMA1.Text = "";
            JK1.Text = "";
            LAHIR1.Value = DateTime.Now;
            PDD1.Text = "";
            AGAMA.Text = "";
            TMK1.Value = DateTime.Now;
            tkk.Value = DateTime.Now;
            USIA1.Text = "";
            ALAMAT1.Text = "";
            KAB1.Text = "";
            DOMI.Text = "";
            KTP1.Text = "";
            IBU1.Text = "";
            hp.Text = "";
            sta.Text = "";
            KET.Text = "";


        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int row = dataGridView1.CurrentCell.RowIndex;
            npkterpilih = dataGridView1.Rows[row].Cells[0].Value.ToString();

            NPK1.Text = dataGridView1.Rows[row].Cells[0].Value.ToString();
            NAMA1.Text = dataGridView1.Rows[row].Cells[1].Value.ToString();
            JK1.Text = dataGridView1.Rows[row].Cells[2].Value.ToString();
            LAHIR1.Value = Convert.ToDateTime(dataGridView1.Rows[row].Cells[3].Value.ToString());
            PDD1.Text = dataGridView1.Rows[row].Cells[4].Value.ToString();
            AGAMA.Text = dataGridView1.Rows[row].Cells[5].Value.ToString();
            TMK1.Value = Convert.ToDateTime(dataGridView1.Rows[row].Cells[6].Value.ToString());
            USIA1.Text = dataGridView1.Rows[row].Cells[7].Value.ToString();
            tkk.Value = Convert.ToDateTime(dataGridView1.Rows[row].Cells[8].Value.ToString());
            ALAMAT1.Text = dataGridView1.Rows[row].Cells[9].Value.ToString();
            KAB1.Text = dataGridView1.Rows[row].Cells[10].Value.ToString();
            DOMI.Text = dataGridView1.Rows[row].Cells[11].Value.ToString();
            KTP1.Text = dataGridView1.Rows[row].Cells[12].Value.ToString();
            IBU1.Text = dataGridView1.Rows[row].Cells[13].Value.ToString();
            hp.Text = dataGridView1.Rows[row].Cells[14].Value.ToString();
            sta.Text = dataGridView1.Rows[row].Cells[15].Value.ToString();
            KET.Text = dataGridView1.Rows[row].Cells[16].Value.ToString();

        }

        private void label16_Click(object sender, EventArgs e)
        {

        }

        private void tOEXCELToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DateTime date = DateTime.Today;
            koneksi.Open();
            cmd1 = new SqlCommand();
            cmd1 = koneksi.CreateCommand();
            cmd1.CommandText = "SELECT BIO.NPK,BIO.NAMA,BIO.JK,BIO.TGLLAHIR,BIO.PDDK,BIO.AGAMA,BIO.TMK,BIO.USIA,BIO.TKK,BIO.BAGIAN,BIO.ALAMAT,BIO.KABUPATEN,BIO.DOMISILI,BIO.KTP,BIO.IBU,BIO.HP,BIO.STATUS,BIO.KETERANGAN FROM [BIO] where tkk is not null order by tkk desc";


            ds = new DataSet();
            da = new SqlDataAdapter(cmd1);
            da.Fill(ds, "DATA kar");
            //export datatable to excel
            DataTable t = ds.Tables[0];
            Workbook book = new Workbook();
            Worksheet sheet = book.Worksheets[0];
            Worksheet shit = book.Worksheets[1];
            sheet.InsertDataTable(t, true, 1, 1);
            //auto collumn
            sheet.AllocatedRange.AutoFitColumns();
            sheet.AllocatedRange.AutoFitRows();

            //style
            sheet.Range["A1:P1"].Style.Color = Color.Gray;
            sheet.Range["A1:P1"].Style.Font.IsBold = true;
            sheet.FreezePanes(2, 1);

            //simpan worksheet dengan save dialog
            System.Windows.Forms.SaveFileDialog saveDlg = new System.Windows.Forms.SaveFileDialog();
                saveDlg.InitialDirectory = @"C:\Users\Acer\Music";
                saveDlg.Filter = "Excel 97-2003 Workbook(*.xls)|*.xls|Excel Workbook(*.xlsx)|*xlsx";
                saveDlg.DefaultExt = ".xlsx";
                saveDlg.FilterIndex = 0;
                saveDlg.RestoreDirectory = true;
                saveDlg.FileName = "Laporan Database Out GMT 12 " + date.ToString(string.Format("dd-MM-yyyy", date)) + " ";
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
            }
        

        private void metroButton7_Click(object sender, EventArgs e)
        {
            this.Close();
            Form6 awal = new Form6();
            awal.ShowDialog();
        }

        private void OUT_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'iWAK.BIO' table. You can move, or remove it, as needed.
          
            // TODO: This line of code loads data into the '_JAPER_DATADataSet5.BIO' table. You can move, or remove it, as needed.
            

            //set datagridview coloum width
            dataGridView1.Columns[0].Width = 70;//NPK
            dataGridView1.Columns[1].Width = 200;//NAMA
            dataGridView1.Columns[2].Width = 30;//JK
            dataGridView1.Columns[3].Width = 80;//TGLLAHIR
            dataGridView1.Columns[4].Width = 50;//PDDK
            dataGridView1.Columns[5].Width = 70;//AGAMA
            dataGridView1.Columns[6].Width = 80;//TMK
            dataGridView1.Columns[7].Width = 200;//USIA
            dataGridView1.Columns[8].Width = 80;//TKK
            dataGridView1.Columns[9].Width = 200;//ALAMAT
            dataGridView1.Columns[10].Width = 100;//KABU
            dataGridView1.Columns[11].Width = 130;//DOMIS
            dataGridView1.Columns[12].Width = 130;//KTP
            dataGridView1.Columns[13].Width = 100;//IBU
            dataGridView1.Columns[14].Width = 100;//HP
            dataGridView1.Columns[15].Width = 70;//STATUS
            dataGridView1.Columns[16].Width = 100;//KETERANGAN

            dateTimePicker2.Enabled = false;
            caritxt.Enabled = false;

        }

        private void caribtn_Click(object sender, EventArgs e)
        {
            if (cbcari.SelectedItem.ToString() == "NAMA")
            {
                cmd1 = new SqlCommand();
                cmd1 = koneksi.CreateCommand();
                cmd1.CommandText = "SELECT BIO.NPK,BIO.NAMA,BIO.JK,BIO.TGLLAHIR,BIO.PDDK,BIO.AGAMA,BIO.TMK,BIO.USIA,BIO.TKK,BIO.ALAMAT,BIO.KABUPATEN,BIO.DOMISILI,BIO.KTP,BIO.IBU,BIO.HP,BIO.STATUS,BIO.KETERANGAN FROM [BIO] WHERE nama like '" + caritxt.Text + "%' AND TKK IS NOT NULL";
                ds = new DataSet();
                da = new SqlDataAdapter(cmd1);
                da.Fill(ds, "DATA IJIN");
                dataGridView1.DataSource = ds;
                dataGridView1.DataMember = "DATA IJIN";

                dataGridView1.Refresh();
                koneksi.Close();
                //Menampilkan Jumlah Data dari DataGridView
                int numRows = dataGridView1.Rows.Count - 1;
                label9.Text = numRows.ToString();
            }
            else if (cbcari.SelectedItem.ToString() == "NPK")
            {
                cmd1 = new SqlCommand();
                cmd1 = koneksi.CreateCommand();
                cmd1.CommandText = "SELECT BIO.NPK,BIO.NAMA,BIO.JK,BIO.TGLLAHIR,BIO.PDDK,BIO.AGAMA,BIO.TMK,BIO.USIA,BIO.TKK,BIO.ALAMAT,BIO.KABUPATEN,BIO.DOMISILI,BIO.KTP,BIO.IBU,BIO.HP,BIO.STATUS,BIO.KETERANGAN FROM [BIO] WHERE NPK like '" + caritxt.Text + "%' AND TKK IS NOT NULL";
                ds = new DataSet();
                da = new SqlDataAdapter(cmd1);
                da.Fill(ds, "DATA IJIN");
                dataGridView1.DataSource = ds;
                dataGridView1.DataMember = "DATA IJIN";

                dataGridView1.Refresh();
                koneksi.Close();
                //Menampilkan Jumlah Data dari DataGridView
                int numRows = dataGridView1.Rows.Count - 1;
                label9.Text = numRows.ToString();
            }
            else if (cbcari.SelectedItem.ToString() == "KABUPATEN")
            {
                cmd1 = new SqlCommand();
                cmd1 = koneksi.CreateCommand();
                cmd1.CommandText = "SELECT BIO.NPK,BIO.NAMA,BIO.JK,BIO.TGLLAHIR,BIO.PDDK,BIO.AGAMA,BIO.TMK,BIO.USIA,BIO.TKK,BIO.ALAMAT,BIO.KABUPATEN,BIO.DOMISILI,BIO.KTP,BIO.IBU,BIO.HP,BIO.STATUS,BIO.KETERANGAN FROM [BIO] WHERE KABUPATEN like '" + caritxt.Text + "%' AND TKK IS NOT NULL";
                ds = new DataSet();
                da = new SqlDataAdapter(cmd1);
                da.Fill(ds, "DATA IJIN");
                dataGridView1.DataSource = ds;
                dataGridView1.DataMember = "DATA IJIN";

                dataGridView1.Refresh();
                koneksi.Close();
                //Menampilkan Jumlah Data dari DataGridView
                int numRows = dataGridView1.Rows.Count - 1;
                label9.Text = numRows.ToString();
            }
            else if (cbcari.SelectedItem.ToString() == "DOMISILI")
            {
                cmd1 = new SqlCommand();
                cmd1 = koneksi.CreateCommand();
                cmd1.CommandText = "SELECT BIO.NPK,BIO.NAMA,BIO.JK,BIO.TGLLAHIR,BIO.PDDK,BIO.AGAMA,BIO.TMK,BIO.USIA,BIO.TKK,BIO.ALAMAT,BIO.KABUPATEN,BIO.DOMISILI,BIO.KTP,BIO.IBU,BIO.HP,BIO.STATUS,BIO.KETERANGAN FROM [BIO] WHERE DOMISILI like '" + caritxt.Text + "%' AND TKK IS NOT NULL";
                ds = new DataSet();
                da = new SqlDataAdapter(cmd1);
                da.Fill(ds, "DATA IJIN");
                dataGridView1.DataSource = ds;
                dataGridView1.DataMember = "DATA IJIN";

                dataGridView1.Refresh();
                koneksi.Close();
                //Menampilkan Jumlah Data dari DataGridView
                int numRows = dataGridView1.Rows.Count - 1;
                label9.Text = numRows.ToString();
            }
            else if (cbcari.SelectedItem.ToString() == "ALAMAT")
            {
                cmd1 = new SqlCommand();
                cmd1 = koneksi.CreateCommand();
                cmd1.CommandText = "SELECT BIO.NPK,BIO.NAMA,BIO.JK,BIO.TGLLAHIR,BIO.PDDK,BIO.AGAMA,BIO.TMK,BIO.USIA,BIO.TKK,BIO.ALAMAT,BIO.KABUPATEN,BIO.DOMISILI,BIO.KTP,BIO.IBU,BIO.HP,BIO.STATUS,BIO.KETERANGAN FROM [BIO] WHERE ALAMAT like '" + caritxt.Text + "%' AND TKK IS NOT NULL";
                ds = new DataSet();
                da = new SqlDataAdapter(cmd1);
                da.Fill(ds, "DATA IJIN");
                dataGridView1.DataSource = ds;
                dataGridView1.DataMember = "DATA IJIN";

                dataGridView1.Refresh();
                koneksi.Close();
                //Menampilkan Jumlah Data dari DataGridView
                int numRows = dataGridView1.Rows.Count - 1;
                label9.Text = numRows.ToString();
            }
            else if (cbcari.SelectedItem.ToString() == "TKK")
            {
                dateTimePicker2.Enabled = true;
                string today = dateTimePicker2.Value.ToString(string.Format("yyyy-MM-dd"));
                cmd1 = new SqlCommand();
                cmd1 = koneksi.CreateCommand();
                cmd1.CommandText = "SELECT BIO.NPK,BIO.NAMA,BIO.JK,BIO.TGLLAHIR,BIO.PDDK,BIO.AGAMA,BIO.TMK,BIO.USIA,BIO.TKK,BIO.ALAMAT,BIO.KABUPATEN,BIO.DOMISILI,BIO.KTP,BIO.IBU,BIO.HP,BIO.STATUS,BIO.KETERANGAN FROM [BIO] WHERE TKK like '" + today + "%' AND TKK IS NOT NULL";
                ds = new DataSet();
                da = new SqlDataAdapter(cmd1);
                da.Fill(ds, "DATA IJIN");
                dataGridView1.DataSource = ds;
                dataGridView1.DataMember = "DATA IJIN";

                dataGridView1.Refresh();
                koneksi.Close();
                //Menampilkan Jumlah Data dari DataGridView
                int numRows = dataGridView1.Rows.Count - 1;
                label9.Text = numRows.ToString();
            }
            else
            {
                cmd1 = new SqlCommand();
                cmd1 = koneksi.CreateCommand();
                cmd1.CommandText = "SELECT BIO.NPK,BIO.NAMA,BIO.JK,BIO.TGLLAHIR,BIO.PDDK,BIO.AGAMA,BIO.TMK,BIO.USIA,BIO.TKK,BIO.ALAMAT,BIO.KABUPATEN,BIO.DOMISILI,BIO.KTP,BIO.IBU,BIO.HP,BIO.STATUS,BIO.KETERANGAN FROM [BIO] WHERE KTP like '" + caritxt.Text + "%' AND TKK IS NOT NULL";
                ds = new DataSet();
                da = new SqlDataAdapter(cmd1);
                da.Fill(ds, "DATA IJIN");
                dataGridView1.DataSource = ds;
                dataGridView1.DataMember = "DATA IJIN";

                dataGridView1.Refresh();
                koneksi.Close();
                //Menampilkan Jumlah Data dari DataGridView
                int numRows = dataGridView1.Rows.Count - 1;
                label9.Text = numRows.ToString();
            }
        }

        private void cbcari_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbcari.SelectedItem.ToString() == "TKK")
            {
                dateTimePicker2.Enabled = true;
                caritxt.Enabled = false;

            }
            else
            {
                caritxt.Enabled = true;
                dateTimePicker2.Enabled = false;
            }
        }

        private void metroButton2_Click(object sender, EventArgs e)
        {
            update();
        }

        private void metroButton4_Click(object sender, EventArgs e)
        {
            reset();
        }

        private void hapus()
        {
            koneksi.Open();
            cmd1 = new SqlCommand();
            cmd1 = koneksi.CreateCommand();
            cmd1.CommandText = "DELETE FROM [BIO] WHERE NPK='" + NPK1.Text + "'";

            try
            {

                //adapter = new OdbcDataAdapter(cmd);
                SqlDataAdapter da = new SqlDataAdapter(cmd1);
                da.DeleteCommand = koneksi.CreateCommand();



                //PROMPT FOR CONFIRMATION BEFORE DELETING
                if (MessageBox.Show(@"Are you sure to permanently delete this?", @"DELETE", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
                {
                    if (cmd1.ExecuteNonQuery() > 0)
                    {

                        MessageBox.Show(@"Successfully deleted");

                    }
                }
                koneksi.Close();
                // gridGroupingControl1.DataSource = null;
                DataTable dt = new DataTable();
                dt.Rows.Clear();
                tampildata();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                koneksi.Close();
            }
        }


        private void metroButton3_Click(object sender, EventArgs e)
        {
            hapus();
        }

        private void metroButton5_Click(object sender, EventArgs e)
        {
            tampildata();
            reset();
        }

        private void tkk_ValueChanged(object sender, EventArgs e)
        {

        }
    }
}
