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
    public partial class Form6 : MetroForm
    {



        //private OdbcDataAdapter adapter;
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
       




        public string npkterpilih;

        //private OdbcDataAdapter adapter;

        public Form6()
        {

            InitializeComponent();
            tampildata();
            this.BindDataGridView();



        }
        //akses database pc1
        const string connString = "Data Source=ANNNIIXX;Initial Catalog=JAPER-DATA;integrated security = true";
        SqlConnection koneksi = new SqlConnection("Data Source=ANNNIIXX;Initial Catalog=JAPER-DATA;integrated security = true");
        //akses database pc2
        //const string connString2 = "Data Source=ANNNIIXX;Initial Catalog=JAPER-DATA;integrated security = true";
        //SqlConnection koneksi2 = new SqlConnection("Data Source=ANNNIIXX;Initial Catalog=JAPER-DATA;integrated security = true");
        private DataView myView;

        private DataTable dataTable = new DataTable();



        private void departe()
        {
            {
                //SQL STATEMENT
                koneksi.Open();
                cmd1 = new SqlCommand();
                cmd1 = koneksi.CreateCommand();
                cmd1.CommandText = "SELECT [ID_DEPT],([DEPARTEMENT] +' - '+ [SECTION]) AS DISPLAY,[LEADER] FROM [dbo].[DEPT]";
                //cmd = new OdbcCommand(sql, con);

                try
                {

                    SqlDataAdapter da = new SqlDataAdapter(cmd1);
                    da.Fill(deptdt);
                    BAG.DataSource = deptdt;
                    BAG.DisplayMember = "DISPLAY";
                    BAG.ValueMember = "ID_DEPT";
                    

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);

                    
                }
                koneksi.Close();

            }
        }

        private void tampildata()
        {

            //SQL STATEMENT
            koneksi.Open();
            cmd1 = new SqlCommand();
            cmd1 = koneksi.CreateCommand();
            cmd1.CommandText = "SELECT top 20 BIO.NPK,BIO.NAMA,BIO.JK,BIO.TGLLAHIR,BIO.PDDK,BIO.AGAMA,BIO.TMK,BIO.USIA,BIO.ALAMAT,BIO.KABUPATEN,BIO.DOMISILI,BIO.KTP,BIO.IBU,BIO.HP,BIO.STATUS,BIO.KETERANGAN FROM [BIO] where tkk is null order by npk asc";
            //cmd = new OdbcCommand(sql, connString);

            ds = new DataSet();
            da = new SqlDataAdapter(cmd1);
            da.Fill(ds, "DATA kar");
            dataGridView1.DataSource = ds;
            dataGridView1.DataMember = "DATA kar";



            dataGridView1.Refresh();

            cmd1 = new SqlCommand();
            cmd1 = koneksi.CreateCommand();
            cmd1.CommandText = "SELECT COUNT(*) As count FROM [BIO] where TKK is null";

            ds = new DataSet();
            da = new SqlDataAdapter(cmd1);
            da.Fill(ds);
            label9.Text = ds.Tables[0].Rows[0]["count"].ToString();
            koneksi.Close();

            koneksi.Open();
            cmd1 = new SqlCommand();
            cmd1 = koneksi.CreateCommand();
            cmd1.CommandText = "SELECT COUNT(*) As pe FROM [BIO] where jk='P' and TKK is null";
            ds = new DataSet();
            da = new SqlDataAdapter(cmd1);
            da.Fill(ds);
            lk.Text = ds.Tables[0].Rows[0]["pe"].ToString();
            koneksi.Close();

            koneksi.Open();
            cmd1 = new SqlCommand();
            cmd1 = koneksi.CreateCommand();
            cmd1.CommandText = "SELECT COUNT(*) As la FROM [BIO] where jk='L' and TKK is null";
            ds = new DataSet();
            da = new SqlDataAdapter(cmd1);
            da.Fill(ds);
            perem.Text = ds.Tables[0].Rows[0]["la"].ToString();
            koneksi.Close();


        }

        private void tambah()
        {
            //SQL STMT

            koneksi.Open();
            cmd1 = new SqlCommand();
            cmd1 = koneksi.CreateCommand();
            cmd1.CommandText = "INSERT INTO BIO (NPK,NAMA,JK,TGLLAHIR,PDDK,AGAMA,TMK,USIA,ALAMAT,KABUPATEN,DOMISILI,KTP,IBU,HP,STATUS,KETERANGAN) VALUES (@NPK2,@NAMA2,@JK2,@TGLLAHIR2,@PDDK2,@AGAMA2,@TMK2,@USIA2,@ALAMAT2,@KABUPATEN2,@DOMISILI2,@KTP2,@IBU2,@HP2,@STATUS2,@KETERANGAN2)";
            //cmd = new OdbcCommand(sql, conn);

            //ADD PARAMS
            cmd1.Parameters.AddWithValue("@NPK2", NPK1.Text);
            cmd1.Parameters.AddWithValue("@NAMA2", NAMA1.Text);
            cmd1.Parameters.AddWithValue("@JK2", JK1.Text);
            cmd1.Parameters.AddWithValue("@TGLLAHIR2", LAHIR1.Value);
            cmd1.Parameters.AddWithValue("@PDDK2", PDD1.Text);
            cmd1.Parameters.AddWithValue("@AGAMA2", AGAMA.Text);
            cmd1.Parameters.AddWithValue("@TMK2", TMK1.Value);
            cmd1.Parameters.AddWithValue("@USIA2", USIA1.Text);
            cmd1.Parameters.AddWithValue("@ALAMAT2", ALAMAT1.Text);
            cmd1.Parameters.AddWithValue("@KABUPATEN2", KAB1.Text);
            cmd1.Parameters.AddWithValue("@DOMISILI2", DOMI.Text);
            cmd1.Parameters.AddWithValue("@KTP2", KTP1.Text);
            cmd1.Parameters.AddWithValue("@IBU2", IBU1.Text);
            cmd1.Parameters.AddWithValue("@HP2", hp.Text);
            cmd1.Parameters.AddWithValue("@STATUS2", sta.Text);
            cmd1.Parameters.AddWithValue("@KETERANGAN2", KET.Text);

            //OPEN CON AND EXEC INSERT
 
            cmd1.ExecuteNonQuery();
            MessageBox.Show("Berhasil Menambahkan Data", "Informasi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            tambah_biodata();
            tampildata();
            reset();
            koneksi.Close();
        }

        private void tambah_biodata()
        {
            koneksi.Open();
            //barcode();
            cmd1 = new SqlCommand();
            cmd1 = koneksi.CreateCommand();
            //barcode();
            cmd1.CommandText = "INSERT INTO [BIODATA](NPK,NAMA_KARYAWAN,BAG,ID_DEPT,JENIS_KEL,BARCODE,SECTION) VALUES(@NPK2,@NAMA2,'OPERATOR',@DEPT,@JK2,@BARC,'GARMENT 12')";
            //cmd = new OdbcCommand(sql, conn);


            //ADD PARAMS
            cmd1.Parameters.AddWithValue("@NPK2", NPK1.Text);
            cmd1.Parameters.AddWithValue("@NAMA2", NAMA1.Text);
            cmd1.Parameters.AddWithValue("@JK2", JK1.Text);
            cmd1.Parameters.AddWithValue("@DEPT", BAG.SelectedValue.ToString());
            cmd1.Parameters.AddWithValue("@BARC", BARC.Text);
            

            //OPEN CON AND EXEC INSERT

            cmd1.ExecuteNonQuery();
            MessageBox.Show("Berhasil Menambahkan Data Ke Database 2", "Informasi", MessageBoxButtons.OK, MessageBoxIcon.Information);
           
            tampildata();
            reset();
            koneksi.Close();
        }

        private void tambah_keluar()
        {
            //SQL STMT

            koneksi.Open();
            cmd1 = new SqlCommand();
            cmd1 = koneksi.CreateCommand();
            cmd1.CommandText = "INSERT INTO [BIODATA_KELUAR] (NPK,NAMA_KARYAWAN,BAG,ID_DEPT,JENIS_KEL,BARCODE,SECTION) SELECT NPK, NAMA_KARYAWAN, BAG, ID_DEPT, JENIS_KEL, BARCODE, SECTION FROM [BIODATA] WHERE NPK ='" + NPK1.Text + "'";
            //cmd = new OdbcCommand(sql, conn);
            cmd1.ExecuteNonQuery();

            koneksi.Close();
        }

        private void tambah_keluar_cbs()
        {
            //SQL STMT

            koneksi.Open();
            cmd1 = new SqlCommand();
            cmd1 = koneksi.CreateCommand();
            cmd1.CommandText = "INSERT INTO [BIODATA_KELUAR_CBS] (NPK,NAMA_KARYAWAN,BAG,ID_DEPT,JENIS_KEL,BARCODE,SECTION) SELECT NPK, NAMA_KARYAWAN, BAG, ID_DEPT, JENIS_KEL, BARCODE, SECTION FROM [BIODATA] WHERE NPK ='" + NPK1.Text + "'";
            //cmd = new OdbcCommand(sql, conn);
            cmd1.ExecuteNonQuery();

            koneksi.Close();
        }


        private void Form6_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'bIO._BIO' table. You can move, or remove it, as needed.
         
            // TODO: This line of code loads data into the 'bIO_BIO.BIO' table. You can move, or remove it, as needed.

            // TODO: This line of code loads data into the 'bIO_JAPER.BIO' table. You can move, or remove it, as needed.

            // TODO: This line of code loads data into the 'bIODataSet.BIO' table. You can move, or remove it, as needed.

            tampildata();

            //set datagridview coloum width
            dataGridView1.Columns[0].Width = 70;//NPK
            dataGridView1.Columns[1].Width = 200;//NAMA
            dataGridView1.Columns[2].Width = 30;//JK
            dataGridView1.Columns[3].Width = 80;//TGLLAHIR
            dataGridView1.Columns[4].Width = 50;//PDD
            dataGridView1.Columns[5].Width = 70;//AGAMA
            dataGridView1.Columns[6].Width = 80;//TMK
            dataGridView1.Columns[7].Width = 200;//USIA
            dataGridView1.Columns[8].Width = 200;//ALAMAT
            dataGridView1.Columns[9].Width = 100;//KAB
            dataGridView1.Columns[10].Width = 130;//DOMISI
            dataGridView1.Columns[11].Width = 130;//KTP
            dataGridView1.Columns[12].Width = 100;//IBU
            dataGridView1.Columns[13].Width = 100;//HP
            dataGridView1.Columns[14].Width = 70;//STATUS
            dataGridView1.Columns[15].Width = 100;//KETERANGAN

            dateTimePicker2.Enabled = false;
            caritxt.Enabled = false;
            metroButton3.Enabled = false;
            //BAG.Enabled = false;
            npkdata();
            barcode();
            departe();
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void toolStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void dATAIZINToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new IZIN().Show();



        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            tambah();
        }



        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //set datagridview value to fill component
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
            ALAMAT1.Text = dataGridView1.Rows[row].Cells[8].Value.ToString();
            KAB1.Text = dataGridView1.Rows[row].Cells[9].Value.ToString();
            DOMI.Text = dataGridView1.Rows[row].Cells[10].Value.ToString();
            KTP1.Text = dataGridView1.Rows[row].Cells[11].Value.ToString();
            IBU1.Text = dataGridView1.Rows[row].Cells[12].Value.ToString();
            hp.Text = dataGridView1.Rows[row].Cells[13].Value.ToString();
            sta.Text = dataGridView1.Rows[row].Cells[14].Value.ToString();
            KET.Text = dataGridView1.Rows[row].Cells[15].Value.ToString();
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void dataGridView1_Scroll(object sender, ScrollEventArgs e)
        {

        }

        private void TMK1_ValueChanged(object sender, EventArgs e)
        {
            //declare variable from datetimepicker
            DateTime start = Convert.ToDateTime(LAHIR1.Value);
            DateTime end = Convert.ToDateTime(TMK1.Value);

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

                //isi textbox dengan data dari hasil selisih
                USIA1.Text = totalyear.ToString() + " Tahun " + totalmonth.ToString() + " Bulan " + totaldays.ToString() + " Hari";

            }

            else
            {
                //jika tanggal salah maka isi data dibawah ini
                USIA1.Text = "0 Tahun 0 Bulan 0 Hari";
                TMK1.Value = DateTime.Now;
                MessageBox.Show("Harap Atur Tanggal Dengan Benar.", "Informasi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }

        private void hapus_biodata()
        {
            koneksi.Open();
            cmd1 = new SqlCommand();
            cmd1 = new SqlCommand();
            cmd1 = koneksi.CreateCommand();
            cmd1.CommandText = "DELETE FROM [BIODATA] WHERE NPK='" + NPK1.Text + "'";
            //cmd = new OdbcCommand(sql, conn);
            cmd1.ExecuteNonQuery();

            koneksi.Close();
        }

        private void hapus_biodata_cbs()
        {
            koneksi.Open();
            cmd1 = new SqlCommand();
            cmd1 = new SqlCommand();
            cmd1 = koneksi.CreateCommand();
            cmd1.CommandText = "DELETE FROM [BIODATA_CBS] WHERE NPK='" + NPK1.Text + "'";
            //cmd = new OdbcCommand(sql, conn);
            cmd1.ExecuteNonQuery();

            koneksi.Close();
        }

        private void hapus()
        {
            SqlConnection conn = new SqlConnection(connString);
            string sql = "UPDATE [BIO] SET tkk='" + TKK.Value.ToString(string.Format("yyyy-MM-dd")) + "',KETERANGAN='" + KET.Text + "',BAGIAN='" + BAG.Text + "' where npk='" + NPK1.Text + "'";
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
                    tambah_keluar();
                    //tambah_keluar_cbs();
                    hapus_biodata();
                    hapus_biodata_cbs();
                    MessageBox.Show(@"Data Dipindah Ke Karyawan OUT");
                }
                conn.Close();
                
                //REFRESH DATA
                DataTable dt = new DataTable();
                dt.Rows.Clear();
                tampildata();
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
            npkdata();
            NAMA1.Text = "";
            JK1.Text = "";
            LAHIR1.Value = DateTime.Now;
            PDD1.Text = "";
            AGAMA.Text = "";
            TMK1.Value = DateTime.Now;
            USIA1.Text = "";
            ALAMAT1.Text = "";
            KAB1.Text = "";
            DOMI.Text = "";
            KTP1.Text = "";
            IBU1.Text = "";
            hp.Text = "";
            sta.Text = "";
            KET.Text = "";
            BAG.Text = "";


        }

        private void metroButton3_Click(object sender, EventArgs e)
        {
            hapus();

        }

        private void eXPORTXMLToolStripMenuItem_Click(object sender, EventArgs e)
        {


        }

        private void metroButton4_Click(object sender, EventArgs e)
        {
            reset();
        }

        private void update()
        {
            SqlConnection conn = new SqlConnection(connString);
            string sql = "UPDATE [BIO] SET nama='" + NAMA1.Text + "',jk='" + JK1.Text + "',tgllahir='" + LAHIR1.Value.ToString(string.Format("yyyy-MM-dd")) + "',pddk='" + PDD1.Text + "',agama='" + AGAMA.Text + "',tmk='" + TMK1.Value.ToString(string.Format("yyyy-MM-dd")) + "',usia='" + USIA1.Text + "',alamat='" + ALAMAT1.Text + "',kabupaten='" + KAB1.Text + "',domisili='" + DOMI.Text + "',ktp='" + KTP1.Text + "',ibu='" + IBU1.Text + "',hp='" + hp.Text + "',status='" + sta.Text + "',KETERANGAN='" + KET.Text + "' where npk='" + NPK1.Text + "'";
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
                tampildata();
                reset();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                koneksi.Close();
            }
        }

        private void metroButton2_Click(object sender, EventArgs e)
        {
            update();
        }

        private void caribtn_Click(object sender, EventArgs e)
        {
            if (cbcari.SelectedItem.ToString() == "NAMA")
            {
                cmd1 = new SqlCommand();
                cmd1 = koneksi.CreateCommand();
                cmd1.CommandText = "SELECT BIO.NPK,BIO.NAMA,BIO.JK,BIO.TGLLAHIR,BIO.PDDK,BIO.AGAMA,BIO.TMK,BIO.USIA,BIO.ALAMAT,BIO.KABUPATEN,BIO.DOMISILI,BIO.KTP,BIO.IBU,BIO.HP,BIO.STATUS,BIO.KETERANGAN FROM [BIO] WHERE nama like '" + caritxt.Text + "%' AND tkk is null";
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
                cmd1.CommandText = "SELECT BIO.NPK,BIO.NAMA,BIO.JK,BIO.TGLLAHIR,BIO.PDDK,BIO.AGAMA,BIO.TMK,BIO.USIA,BIO.ALAMAT,BIO.KABUPATEN,BIO.DOMISILI,BIO.KTP,BIO.IBU,BIO.HP,BIO.STATUS,BIO.KETERANGAN FROM [BIO] WHERE NPK like '" + caritxt.Text + "%' AND tkk is null";
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
            else if (cbcari.SelectedItem.ToString() == "TANGGAL LAHIR")
            {
                dateTimePicker2.Enabled = true;
                string today = dateTimePicker2.Value.ToString(string.Format("yyyy-MM-dd"));
                cmd1 = new SqlCommand();
                cmd1 = koneksi.CreateCommand();
                cmd1.CommandText = "SELECT BIO.NPK,BIO.NAMA,BIO.JK,BIO.TGLLAHIR,BIO.PDDK,BIO.AGAMA,BIO.TMK,BIO.USIA,BIO.ALAMAT,BIO.KABUPATEN,BIO.DOMISILI,BIO.KTP,BIO.IBU,BIO.HP,BIO.STATUS,BIO.KETERANGAN FROM [BIO] WHERE TGLLAHIR like '" + today + "%' AND tkk is null";
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
                cmd1.CommandText = "SELECT BIO.NPK,BIO.NAMA,BIO.JK,BIO.TGLLAHIR,BIO.PDDK,BIO.AGAMA,BIO.TMK,BIO.USIA,BIO.ALAMAT,BIO.KABUPATEN,BIO.DOMISILI,BIO.KTP,BIO.IBU,BIO.HP,BIO.STATUS,BIO.KETERANGAN FROM [BIO] WHERE KABUPATEN like '" + caritxt.Text + "%' AND tkk is null";
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
                cmd1.CommandText = "SELECT BIO.NPK,BIO.NAMA,BIO.JK,BIO.TGLLAHIR,BIO.PDDK,BIO.AGAMA,BIO.TMK,BIO.USIA,BIO.ALAMAT,BIO.KABUPATEN,BIO.DOMISILI,BIO.KTP,BIO.IBU,BIO.HP,BIO.STATUS,BIO.KETERANGAN FROM [BIO] WHERE DOMISILI like '" + caritxt.Text + "%' AND tkk is null";
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
                cmd1.CommandText = "SELECT BIO.NPK,BIO.NAMA,BIO.JK,BIO.TGLLAHIR,BIO.PDDK,BIO.AGAMA,BIO.TMK,BIO.USIA,BIO.ALAMAT,BIO.KABUPATEN,BIO.DOMISILI,BIO.KTP,BIO.IBU,BIO.HP,BIO.STATUS,BIO.KETERANGAN FROM [BIO] WHERE ALAMAT like '" + caritxt.Text + "%' AND tkk is null";
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
            else if (cbcari.SelectedItem.ToString() == "TMK")
            {
                dateTimePicker2.Enabled = true;
                string today = dateTimePicker2.Value.ToString(string.Format("yyyy-MM-dd"));
                cmd1 = new SqlCommand();
                cmd1 = koneksi.CreateCommand();
                cmd1.CommandText = "SELECT BIO.NPK,BIO.NAMA,BIO.JK,BIO.TGLLAHIR,BIO.PDDK,BIO.AGAMA,BIO.TMK,BIO.USIA,BIO.ALAMAT,BIO.KABUPATEN,BIO.DOMISILI,BIO.KTP,BIO.IBU,BIO.HP,BIO.STATUS,BIO.KETERANGAN,BIO.KETERANGAN FROM [BIO] WHERE TMK like '" + today + "%' AND tkk is null";
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
                cmd1.CommandText = "SELECT BIO.NPK,BIO.NAMA,BIO.JK,BIO.TGLLAHIR,BIO.PDDK,BIO.AGAMA,BIO.TMK,BIO.USIA,BIO.ALAMAT,BIO.KABUPATEN,BIO.DOMISILI,BIO.KTP,BIO.IBU,BIO.HP,BIO.STATUS,BIO.KETERANGAN FROM [BIO] WHERE KTP like '" + caritxt.Text + "%' AND tkk is null";
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
            if (cbcari.SelectedItem.ToString() == "TMK")
            {
                dateTimePicker2.Enabled = true;
                caritxt.Enabled = false;

            }
            else if (cbcari.SelectedItem.ToString() == "TANGGAL LAHIR")
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

        private void metroButton7_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        private void npkdata()
        {
            long hitung;
            string urut;


            SqlConnection conn = new SqlConnection(connString);
            conn.Open();
            // Perintah untuk mendapatkan nilai terbesar dari field nomor
            string sql = "select npk from BIO where npk in (select max(npk) from BIO) order by npk desc";
            //con = new OdbcConnection(conString);
            SqlCommand cmd = new SqlCommand(sql, conn);
            rd = cmd.ExecuteReader();
            rd.Read();
            //data ditemukan
            if (rd.HasRows)
            {
                // Menambahkan data dari field nomor
                hitung = Convert.ToInt64(rd[0].ToString().Substring(rd["npk"].ToString().Length - 5, 5)) + 1;
                string joinstr = "00000" + hitung;
                // Mengambil 4 karakter kanan terakhir dari string joinstr lalu di tambahkan dengan string URUT
                urut = "K." + joinstr.Substring(joinstr.Length - 5, 5);
            }
            else
            {
                // Jika tidak ditemukan maka mengisi variable urut dengan YPMB-0001
                urut = "K.00001";
            }

            rd.Close();
            NPK1.Enabled = false;
            NPK1.Text = urut;
            conn.Close();
         
        }
        private void barcode()
        {
            DataTable barcdta = new DataTable();
            SqlConnection conn = new SqlConnection(connString);
            string sql = "select max(barcode) from BIODATA";
            //con = new OdbcConnection(conString);
            SqlCommand cmd = new SqlCommand(sql, conn);
            conn.Open();

            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            //adapter = new OdbcDataAdapter("", conn);

            adapter.Fill(barcdta);
            
            Object o = barcdta.Rows[0][0];
            int br = Convert.ToInt32(o) + 1;
            BARC.Text = Convert.ToString(br);
            conn.Close();
        }

        private void kARYAWANOUTToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new OUT().Show();
            this.Hide();
        }

        private void TKK_ValueChanged(object sender, EventArgs e)
        {
            metroButton3.Enabled = true;
            BAG.Enabled = true;
        }

        private void label17_Click(object sender, EventArgs e)
        {

        }

        private void metroButton5_Click(object sender, EventArgs e)
        {
            tampildata();
            npkdata();
            reset();
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {

        }

        private void KTP1_Click(object sender, EventArgs e)
        {

        }

        private void DOMI_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void KAB1_TextChanged(object sender, EventArgs e)
        {
            if (KAB1.Text == "SUKOHARJO")
            {
                DOMI.Text = "SUKOHARJO";
            }
            else
            {
                DOMI.Text = "LUAR SUKOHARJO";
            }

        }

        private void tOEXCELToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DateTime date = DateTime.Today;

            //menggunakan ExcelEngine
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;

                //Create a workbook with single worksheet
                IWorkbook workbook = application.Workbooks.Create(1);

                IWorksheet worksheet = workbook.Worksheets[0];

                //Import from DataGridView to worksheet
                worksheet.ImportDataGridView(dataGridView1, 1, 1, isImportHeader: true, isImportStyle: true);
                
                //set worksheet agar auto fit
                worksheet.UsedRange.AutofitColumns();
                //worksheet.UsedRange.NumberFormat = "@";
                
                

                //simpan worksheet dengan save dialog
                System.Windows.Forms.SaveFileDialog saveDlg = new System.Windows.Forms.SaveFileDialog();
                saveDlg.InitialDirectory = @"C:\Users\Acer\Music";
                saveDlg.Filter = "Excel 97-2003 Workbook(*.xls)|*.xls|Excel Workbook(*.xlsx)|*xlsx";
                saveDlg.DefaultExt = ".xlsx";
                saveDlg.FilterIndex = 0;
                saveDlg.RestoreDirectory = true;
                saveDlg.FileName = "Laporan Database Aktif GMT 12 " + date.ToString(string.Format("dd-MM-yyyy", date)) + " ";
                if (saveDlg.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveDlg.CheckPathExists)
                {

                    MessageBox.Show("Berhasil di Export");

                    string path = saveDlg.FileName;
                    workbook.SaveAs(path);
                    workbook.Saved = true;


                }
            }
        }



        private void BindDataGridView()
        {
            tampildata();
        }

        

        private void toExToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DateTime date = DateTime.Today;
            //koneksi.Open();
            cmd1 = new SqlCommand();
            cmd1 = koneksi.CreateCommand();
            cmd1.CommandText = "SELECT BIO.NPK,BIO.NAMA,BIO.JK,BIO.TGLLAHIR,BIO.PDDK,BIO.AGAMA,BIO.TMK,BIO.USIA,BIO.ALAMAT,BIO.KABUPATEN,BIO.DOMISILI,BIO.KTP,BIO.IBU,BIO.HP,BIO.STATUS,BIO.KETERANGAN FROM [BIO] where tkk is null order by npk asc";
            //cmd = new OdbcCommand(sql, connString);
            ds = new DataSet();
            da = new SqlDataAdapter(cmd1);
            da.Fill(ds, "DATA kar");
            //export datatable to excel
            DataTable t = ds.Tables[0];
            Workbook book = new Workbook();
            Worksheet sheet = book.Worksheets[0];
            sheet.InsertDataTable(t, true, 1, 1);
            sheet.Name = "AKTIF";
            //auto collumn
            sheet.AllocatedRange.AutoFitColumns();
            sheet.AllocatedRange.AutoFitRows();
            //style
            sheet.Range["A1:P1"].Style.Color = Color.Gray;
            sheet.Range["A1:P1"].Style.Font.IsBold = true;
            sheet.FreezePanes(2, 1);

            cmd2 = new SqlCommand();
            cmd2 = koneksi.CreateCommand();
            cmd2.CommandText = "SELECT BIO.NPK,BIO.NAMA,BIO.JK,BIO.TGLLAHIR,BIO.PDDK,BIO.AGAMA,BIO.TMK,BIO.USIA,BIO.TKK,BIO.BAGIAN,BIO.ALAMAT,BIO.KABUPATEN,BIO.DOMISILI,BIO.KTP,BIO.IBU,BIO.HP,BIO.STATUS,BIO.KETERANGAN FROM [BIO] where tkk is not null order by tkk desc";
            ds2 = new DataSet();
            da2 = new SqlDataAdapter(cmd2);
            da2.Fill(ds2, "DATA out");
            DataTable t2 = ds2.Tables[0];
            
            Worksheet sheet2 = book.Worksheets[1];
            sheet2.InsertDataTable(t2, true, 1, 1);
            sheet2.Name = "OUT";
            //auto collumn
            sheet2.AllocatedRange.AutoFitColumns();
            sheet2.AllocatedRange.AutoFitRows();
            //style
            sheet2.Range["A1:R1"].Style.Color = Color.Gray;
            sheet2.Range["A1:R1"].Style.Font.IsBold = true;
            sheet2.FreezePanes(2, 1);



            //book.SaveToFile("insertTableToExcel.xls");
            //System.Diagnostics.Process.Start("insertTableToExcel.xls");

            System.Windows.Forms.SaveFileDialog saveDlg = new System.Windows.Forms.SaveFileDialog();
            saveDlg.InitialDirectory = @"C:\Users\Acer\Music";
            saveDlg.Filter = "Excel 97-2003 Workbook(*.xls)|*.xls|Excel Workbook(*.xlsx)|*xlsx";
            saveDlg.DefaultExt = ".xlsx";
            saveDlg.FilterIndex = 0;
            saveDlg.RestoreDirectory = true;
            saveDlg.FileName = "Laporan Database GMT 12 " + date.ToString(string.Format("dd-MM-yyyy", date)) + " ";
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

        private void metroButton6_Click(object sender, EventArgs e)
        {
            NPK1.Enabled = true;
        }

        private void rECAPKARYAWANKELUARToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new KAR_OUT().Show();
        }
    }
}
            
        

    

