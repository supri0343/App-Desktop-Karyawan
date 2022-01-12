using MetroFramework.Forms;
using System.Data;
using System.Data.SqlClient;
using Syncfusion.XlsIO;
using System;
using System.Data;
using System.Data.Odbc;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Drawing.Imaging;
using Spire.Xls;

namespace JPT_KARPRO
{

    public partial class IZIN : Form
    {

        public string npkterpilih;
        private SqlCommand cmd1;
        private SqlDataAdapter da;
        private DataSet ds;
        private DataTable deptdt = new DataTable();
        private DataTable dt;




        public IZIN()
        {
            InitializeComponent();
            tampildata();
            departe();






        }

        int rno = 0;
        MemoryStream ms;
        byte[] photo_aray;


        const string connString = "Data Source=ANNNIIXX;Initial Catalog=JAPER-DATA;integrated security = true";
        SqlConnection koneksi = new SqlConnection("Data Source=ANNNIIXX;Initial Catalog=JAPER-DATA;integrated security = true");

        private void tampildata()
        {

            cmd1 = new SqlCommand();
            cmd1 = koneksi.CreateCommand();
            cmd1.CommandText = "select IJIN.NPK, IJIN.NAMA, DEPT.DEPARTEMENT, IJIN.JENIS, IJIN.KET, IJIN.TANGGAL,IJIN.SAMPAI, IJIN.GAMBAR from IJIN INNER JOIN DEPT ON IJIN.ID_DEPT=DEPT.ID_DEPT";
            ds = new DataSet();
            da = new SqlDataAdapter(cmd1);
            da.Fill(ds, "DATA IJINi");
            dataGridView1.DataSource = ds;
            dataGridView1.DataMember = "DATA IJINi";



            dataGridView1.Refresh();
            koneksi.Close();

        }


        private void IZIN_Load(object sender, EventArgs e)
        {
            tampildata();


        }

        private void hapus()
        {
            koneksi.Open();
            cmd1 = new SqlCommand();
            cmd1 = koneksi.CreateCommand();
            cmd1.CommandText = "DELETE FROM [IJIN] WHERE NPK='" + npkterpilih + "'";

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

                // gridGroupingControl1.DataSource = null;
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

        private void update()
        {

            koneksi.Open();
           
            SqlConnection conn = new SqlConnection(connString);
            
            string sql = "UPDATE [IJIN] SET nama='" + namatxt.Text + "',ID_DEPT='" + DEP.SelectedValue + "',jenis='" + jenisCB.Text + "',ket='" + kettxt.Text + "',tanggal='" + dateTimePicker1.Value.ToString(string.Format("yyyy-MM-dd")) + "',sampai='" + dateTimePicker3.Value.ToString(string.Format("yyyy-MM-dd")) + "'  where npk='" + npktxt.Text + "'";
            SqlCommand cmd = new SqlCommand(sql, conn);
            
            cmd1 = new SqlCommand();
            cmd1 = koneksi.CreateCommand();
            //cmd1.Parameters.AddWithValue("@photo", photo_aray);
            //cmd1.Parameters.AddWithValue("@photo", SqlDbType.Binary).Value = photo_aray;
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
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                koneksi.Close();
            }
        }

        void conv_photo()
        {
            //converting photo to binary data
            if (pictureBox1.Image != null)
            {
                //using FileStream:(will not work while updating, if image is not changed)
                //FileStream fs = new FileStream(openFileDialog1.FileName, FileMode.Open, FileAccess.Read);
                //byte[] photo_aray = new byte[fs.Length];
                //fs.Read(photo_aray, 0, photo_aray.Length);  

                //using MemoryStream:
                ms = new MemoryStream();
                pictureBox1.Image.Save(ms, ImageFormat.Jpeg);
                byte[] photo_aray = new byte[ms.Length];
                ms.Position = 0;
                ms.Read(photo_aray, 0, photo_aray.Length);
                cmd1.Parameters.AddWithValue("@photo", photo_aray);
            }
        }


        private void tambah()
        {

            koneksi.Open();
            cmd1 = new SqlCommand();
            cmd1 = koneksi.CreateCommand();
            conv_photo();
            cmd1.CommandText = "insert into IJIN (npk,nama,ID_DEPT,jenis,ket,tanggal,sampai,gambar) values (@npk,@nama,@ID_DEPT,@jenis,@ket,@tanggal,@sampai,@photo)";
            cmd1.Parameters.AddWithValue("@npk", npktxt.Text);
            cmd1.Parameters.AddWithValue("@nama", namatxt.Text);
            cmd1.Parameters.AddWithValue("@ID_DEPT", DEP.SelectedValue.ToString());
            cmd1.Parameters.AddWithValue("@jenis", jenisCB.Text);
            cmd1.Parameters.AddWithValue("@ket", kettxt.Text);
            cmd1.Parameters.AddWithValue("@tanggal", dateTimePicker1.Value);
            cmd1.Parameters.AddWithValue("@sampai", dateTimePicker3.Value);



            cmd1.ExecuteNonQuery();
            MessageBox.Show("Berhasil Menambahkan Data", "Informasi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            koneksi.Close();
            reset();
            tampildata();
        }

        private void departe()
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
                DEP.DataSource = deptdt;
                DEP.DisplayMember = "DISPLAY";
                DEP.ValueMember = "ID_DEPT";
                koneksi.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

                koneksi.Close();
            }

        }




        private void IZIN_Load_1(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'bindingSource1.IJIN' table. You can move, or remove it, as needed.
            dateTimePicker2.Enabled = false;
            tampildata();
            caritxt.Enabled = false;

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView_CellClik(object sender, DataGridViewCellEventArgs e)
        {




        }

        private void button1_Click(object sender, EventArgs e)
        {
            tambah();
            tampildata();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            update();
            tampildata();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            hapus();
            tampildata();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }



        private void toolStripLabel1_Click(object sender, EventArgs e)
        {


        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int row = dataGridView1.CurrentCell.RowIndex;
            npkterpilih = dataGridView1.Rows[row].Cells[0].Value.ToString();
            npktxt.Text = dataGridView1.Rows[row].Cells[0].Value.ToString();
            namatxt.Text = dataGridView1.Rows[row].Cells[1].Value.ToString();
            DEP.Text = dataGridView1.Rows[row].Cells[2].Value.ToString();
            jenisCB.Text = dataGridView1.Rows[row].Cells[3].Value.ToString();
            kettxt.Text = dataGridView1.Rows[row].Cells[4].Value.ToString();
            dateTimePicker1.Value = Convert.ToDateTime(dataGridView1.Rows[row].Cells[5].Value.ToString());
            dateTimePicker3.Value = Convert.ToDateTime(dataGridView1.Rows[row].Cells[6].Value.ToString());
            //photo_aray = (byte[])ds.Tables[0].Rows[rno][6];
            try
            {
                pictureBox1.Image = null;
                if (dataGridView1.Rows[row].Cells[7] != null)
                {

                    Byte[] data = new Byte[0];
                    //data = (Byte[])(dataGridView1.Rows[row].Cells[7].Value);
                    data = (Byte[])dataGridView1.Rows[row].Cells[7].Value;
                    MemoryStream mem = new MemoryStream(data);
                    pictureBox1.Image = Image.FromStream(mem);
                }
                else
                {
                    return;
                }
            }
            catch(Exception)
            {
                pictureBox1.Image = null;
                return;
            }

        
       


        }

        private void reset()
        {
            npktxt.Text = "";
            namatxt.Text = "";
            jenisCB.Text = "";
            kettxt.Text = "";
            dateTimePicker1.Value = DateTime.Now;
            dateTimePicker3.Value = DateTime.Now;
            pictureBox1.Image = null;
        }

        private void button4_Click(object sender, EventArgs e)
        {

            reset();
            tampildata();


        }




        private void eXPORTToolStripMenuItem_Click(object sender, EventArgs e)

        {

            DateTime date = DateTime.Today;
            koneksi.Open();

            cmd1 = new SqlCommand();
            cmd1 = koneksi.CreateCommand();
            cmd1.CommandText = "select IJIN.NPK, IJIN.NAMA, DEPT.DEPARTEMENT, IJIN.JENIS, IJIN.KET, IJIN.TANGGAL, IJIN.SAMPAI from IJIN INNER JOIN DEPT ON IJIN.ID_DEPT=DEPT.ID_DEPT";
            ds = new DataSet();
            da = new SqlDataAdapter(cmd1);
            da.Fill(ds, "DATA IJINi");
            DataTable t = ds.Tables[0];
            Workbook book = new Workbook();
            Worksheet sheet = book.Worksheets[0];
            sheet.InsertDataTable(t, true, 1, 1);
            sheet.Name = "DATA IZIN";
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
            saveDlg.FileName = "Laporan IZIN GMT 12 " + date.ToString(string.Format("dd-MM-yyyy", date)) + " ";
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

                koneksi.Close();
            }
        }
  


   

        private void caribtn_Click(object sender, EventArgs e)
        {
            
        }

        private void caritxt_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            cmd1 = new SqlCommand();
            cmd1 = koneksi.CreateCommand();
            cmd1.CommandText = "select IJIN.NPK, IJIN.NAMA, DEPT.DEPARTEMENT, IJIN.JENIS, IJIN.KET, IJIN.TANGGAL from IJIN INNER JOIN DEPT ON IJIN.ID_DEPT=DEPT.ID_DEPT WHERE nama like '" + caritxt.Text + "%'";
            ds = new DataSet();
            da = new SqlDataAdapter(cmd1);
            da.Fill(ds, "DATA IJIN");
            dataGridView1.DataSource = ds;
            dataGridView1.DataMember = "DATA IJIN";



            dataGridView1.Refresh();
            koneksi.Close();
        }

      

        private void dataGridView1_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            if (cbcari.SelectedItem.ToString() == "TANGGAL")
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

        private void btncari_Click(object sender, EventArgs e)
        {
            if (cbcari.SelectedItem.ToString() == "NAMA")
            {
                cmd1 = new SqlCommand();
                cmd1 = koneksi.CreateCommand();
                cmd1.CommandText = "select IJIN.NPK, IJIN.NAMA, DEPT.DEPARTEMENT, IJIN.JENIS, IJIN.KET, IJIN.TANGGAL, IJIN.SAMPAI, IJIN.GAMBAR from IJIN INNER JOIN DEPT ON IJIN.ID_DEPT=DEPT.ID_DEPT WHERE nama like '" + caritxt.Text + "%'";
                ds = new DataSet();
                da = new SqlDataAdapter(cmd1);
                da.Fill(ds, "DATA IJIN");
                dataGridView1.DataSource = ds;
                dataGridView1.DataMember = "DATA IJIN";

                dataGridView1.Refresh();
                koneksi.Close();
            }
            else if(cbcari.SelectedItem.ToString() == "JENIS IZIN")
            {
                cmd1 = new SqlCommand();
                cmd1 = koneksi.CreateCommand();
                cmd1.CommandText = "select IJIN.NPK, IJIN.NAMA, DEPT.DEPARTEMENT, IJIN.JENIS, IJIN.KET, IJIN.TANGGAL, IJIN.SAMPAI, IJIN.GAMBAR from IJIN INNER JOIN DEPT ON IJIN.ID_DEPT=DEPT.ID_DEPT WHERE JENIS like '" + caritxt.Text + "%'";
                ds = new DataSet();
                da = new SqlDataAdapter(cmd1);
                da.Fill(ds, "DATA IJIN");
                dataGridView1.DataSource = ds;
                dataGridView1.DataMember = "DATA IJIN";

                dataGridView1.Refresh();
                koneksi.Close();
            }
            else if(cbcari.SelectedItem.ToString() == "TANGGAL")
            {
                dateTimePicker2.Enabled = true;
                string today = dateTimePicker2.Value.ToString(string.Format("yyyy-MM-dd"));
                cmd1 = new SqlCommand();
                cmd1 = koneksi.CreateCommand();
                cmd1.CommandText = "select IJIN.NPK, IJIN.NAMA, DEPT.DEPARTEMENT, IJIN.JENIS, IJIN.KET, IJIN.TANGGAL, IJIN.SAMPAI, IJIN.GAMBAR from IJIN INNER JOIN DEPT ON IJIN.ID_DEPT=DEPT.ID_DEPT WHERE tanggal like '" + today + "%'";
                ds = new DataSet();
                da = new SqlDataAdapter(cmd1);
                da.Fill(ds, "DATA IJIN");
                dataGridView1.DataSource = ds;
                dataGridView1.DataMember = "DATA IJIN";



                dataGridView1.Refresh();
                koneksi.Close();
            }
            else
            {
                cmd1 = new SqlCommand();
                cmd1 = koneksi.CreateCommand();
                cmd1.CommandText = "select IJIN.NPK, IJIN.NAMA, DEPT.DEPARTEMENT, IJIN.JENIS, IJIN.KET, IJIN.TANGGAL, IJIN.SAMPAI, IJIN.GAMBAR from IJIN INNER JOIN DEPT ON IJIN.ID_DEPT=DEPT.ID_DEPT WHERE npk like '" + caritxt.Text + "%'";
                ds = new DataSet();
                da = new SqlDataAdapter(cmd1);
                da.Fill(ds, "DATA IJIN");
                dataGridView1.DataSource = ds;
                dataGridView1.DataMember = "DATA IJIN";

                dataGridView1.Refresh();
                koneksi.Close();
            }
        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "jpeg|*.jpg|bmp|*.bmp|all files|*.*";
            DialogResult res = openFileDialog1.ShowDialog();
            if (res == DialogResult.OK)
            {
                pictureBox1.Image = Image.FromFile(openFileDialog1.FileName);
            }
        }

        private void openFileDialog1_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {

        }

        private void IZIN_MouseEnter(object sender, EventArgs e)
        {

        }

        private void pictureBox1_MouseEnter(object sender, EventArgs e)
        {
          
        }

        private void pictureBox1_MouseLeave(object sender, EventArgs e)
        {
            
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            pictureBox1.Size = new Size(650, 650);
            //pictureBox1.SizeMode = PictureBoxSizeMode.Zoom;
            button4.Visible = false;
            button5.Visible = false;
            button7.Visible = false;

        }

        private void pictureBox1_DoubleClick(object sender, EventArgs e)
        {
            pictureBox1.Size = new Size(237, 196);
            button5.Visible = true;
            button4.Visible = true;
            button7.Visible = true;
            pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void jenisCB_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void DEP_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void kettxt_TextChanged(object sender, EventArgs e)
        {

        }

        private void namatxt_TextChanged(object sender, EventArgs e)
        {

        }

        private void npktxt_TextChanged(object sender, EventArgs e)
        {

        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            
            koneksi.Open();
            conv_photo();

            cmd1 = new SqlCommand();
            cmd1 = koneksi.CreateCommand();
            cmd1.CommandText = "UPDATE [IJIN] SET gambar=null where npk='" + npktxt.Text + "' ";
            //cmd = new OdbcCommand(sql, conn);
            cmd1.ExecuteNonQuery();
            pictureBox1.Image = null;
            tampildata();
            koneksi.Close();
        }


        private void myPrintDocument2_PrintPage(System.Object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            Bitmap myBitmap1 = new Bitmap(pictureBox1.Width, pictureBox1.Height);
            pictureBox1.DrawToBitmap(myBitmap1, new Rectangle(0, 0, pictureBox1.Width, pictureBox1.Height));
            e.Graphics.DrawImage(myBitmap1, 0, 0);
            myBitmap1.Dispose();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            System.Drawing.Printing.PrintDocument myPrintDocument1 = new System.Drawing.Printing.PrintDocument();
            PrintDialog myPrinDialog1 = new PrintDialog();
            myPrintDocument1.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(myPrintDocument2_PrintPage);
            myPrinDialog1.Document = myPrintDocument1;
            if (myPrinDialog1.ShowDialog() == DialogResult.OK)
            {
                myPrintDocument1.Print();
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            koneksi.Open();
            

            cmd1 = new SqlCommand();
            cmd1 = koneksi.CreateCommand();
            conv_photo();
            cmd1.CommandText = "UPDATE [IJIN] SET gambar=@photo where npk='" + npktxt.Text + "' ";
            //cmd = new OdbcCommand(sql, conn);
           
            //cmd1.Parameters.AddWithValue("@photo", photo_aray);
            cmd1.ExecuteNonQuery();
            pictureBox1.Image = null;
            tampildata();
            koneksi.Close();
        }
    }   
}
