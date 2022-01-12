using MetroFramework.Forms;
using System.Data;
using System.Data.SqlClient;
using Syncfusion.Grouping;
using Syncfusion.GroupingGridExcelConverter;
using Syncfusion.Windows.Forms.Grid;
using Syncfusion.Windows.Forms.Grid.Grouping;
using Syncfusion.XlsIO;
using System;
using System.Data;
using System.Data.Odbc;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace JPT_KARPRO
{
    public partial class Form1 : MetroForm
    {
        //private const string conString = "Dsn=SERVERPC;trusted_connection=Yes;app=Microsoft® Visual Studio® 2015;wsid=DESKTOP-4286IBB;database=JAPER-DATA";
        //private OdbcConnection con = new OdbcConnection(conString);

        // readonly OleDbConnection con = new OleDbConnection(Properties.Settings.Default.PAYROLL_BMGConnectionString);
//        private OdbcCommand cmd;

        private SqlCommand cmd;

        //private OdbcDataAdapter adapter;
        private DataTable dt = new DataTable();
        private DataTable deptdt = new DataTable();

        public Form1()
        {
            InitializeComponent();
            //gridGroupingControl1.ColumnCount = 7;
            // gridGroupingControl1.Columns[0].Name = "NPK";
            //  gridGroupingControl1.Columns[1].Name = "NAMA_KARYAWAN";
            // gridGroupingControl1.Columns[2].Name = "BAG";
            // gridGroupingControl1.Columns[3].Name = "DEPARTEMENT";
            // gridGroupingControl1.Columns[4].Name = "JENIS_KEL";
            //  gridGroupingControl1.Columns[5].Name = "BARCODE";
            //   gridGroupingControl1.Columns[6].Name = "SECTION";

            //    gridGroupingControl1.Columns[0].HeaderText = "NPK";
            //   gridGroupingControl1.Columns[1].HeaderText = "NAMA";
            //  gridGroupingControl1.Columns[2].HeaderText = "BAGIAN";
            //   gridGroupingControl1.Columns[3].HeaderText = "DEPT";
            //  gridGroupingControl1.Columns[4].HeaderText = "JK";
            //    gridGroupingControl1.Columns[5].HeaderText = "BARCODE";
            //    gridGroupingControl1.Columns[6].HeaderText = "GM";

            //gridGroupingControl1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            //SELECTION MODE
            //gridGroupingControl1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            //gridGroupingControl1.MultiSelect = false;
            this.gridGroupingControl1.TableOptions.AllowSelection = GridSelectionFlags.Row;
            // this.gridGroupingControl1.TableModel.ColWidths.ResizeToFit(GridRangeInfo.Table(), GridResizeToFitOptions.IncludeHeaders);
            //this.gridGroupingControl1.TableModel.ColWidths.ResizeToFit(GridRangeInfo.Cells(4, 5, 6, 8));
        }

        const string connString = "Data Source=192.168.1.56,1433;Network Library=DBMSSOCN;Initial Catalog=JAPER-DATA;Integrated Security=True";

        private DataView myView;

        private DataTable dataTable = new DataTable();
        

        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the '_JAPER_DATADataSet.DEPT' table. You can move, or remove it, as needed.
            //this.dEPTTableAdapter.Fill(this._JAPER_DATADataSet.DEPT);
            label1.Text = DateTime.Now.ToLongTimeString();
            timer1.Interval = 1000;
            timer1.Enabled = true;
            retrieve();
            departe();
            barcode();
            BARC.Enabled = false;
           
        }

        private void departe()
        {
            //SQL STATEMENT
            SqlConnection conn = new SqlConnection(connString);
            string sql = "SELECT [ID_DEPT],([DEPARTEMENT] +' - '+ [SECTION]) AS DISPLAY,[LEADER] FROM [dbo].[DEPT]";
            //cmd = new OdbcCommand(sql, con);
            SqlCommand cmd = new SqlCommand(sql, conn);
            try
            {
                conn.Open();
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                adapter.Fill(deptdt);
                DEP.DataSource = deptdt;
                DEP.DisplayMember = "DISPLAY";
                DEP.ValueMember = "ID_DEPT";
                conn.Close();
               
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                conn.Close();
            }
        }

        /*public void PullData()
        {
            const string connString = "Data Source=192.168.1.56,1433;Network Library=DBMSSOCN;Initial Catalog=JAPER-DATA;Integrated Security=True";
            string query = "select * from table";

            SqlConnection conn = new SqlConnection(connString);
            SqlCommand cmd = new SqlCommand(query, conn);
            conn.Open();

            // create data adapter
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            // this will query your database and return the result to your datatable
            da.Fill(dataTable);
            conn.Close();
            da.Dispose();
        }*/

        private void retrieve()
        {
            //SQL STATEMENT
            SqlConnection conn = new SqlConnection(connString);
            string sql = "SELECT BIODATA.NPK,BIODATA.NAMA_KARYAWAN,BIODATA.BAG,DEPT.DEPARTEMENT,BIODATA.JENIS_KEL,BIODATA.BARCODE,BIODATA.SECTION FROM [dbo].[BIODATA] INNER JOIN DEPT ON BIODATA.ID_DEPT=DEPT.ID_DEPT";
            //cmd = new OdbcCommand(sql, connString);
            SqlCommand cmd = new SqlCommand(sql, conn);
            try
            {
                conn.Open();
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                adapter.Fill(dt);
                // myView = dt.DefaultView;
                gridGroupingControl1.DataSource = dt;
                // gridGroupingControl1.DataSource = myView;
                //LOOP THROUGH DATATABLE
                //foreach (DataRow row in dt.Rows)
                //{
                // populate(row[0].ToString(), row[1].ToString(), row[2].ToString(), row[3].ToString(), row[4].ToString(), row[5].ToString(), row[6].ToString());
                // }

                conn.Close();
                //CLEAR DATATABLE
                //dt.Rows.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                conn.Close();
            }
        }
       

        private void rETRIEVEToolStripMenuItem_Click(object sender, EventArgs e)
        {
            retrieve();
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
            conn.Close();
            Object o = barcdta.Rows[0][0];
            int br = Convert.ToInt32(o) + 1;
            BARC.Text = Convert.ToString(br);
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
        }

        private void BARC_Enter(object sender, EventArgs e)
        {
            if (BARC.Text == "" && BARC.Enabled == true)
            {
                barcode();
                BARC.Enabled = false;
            }
            else
            {
                BARC.Enabled = true;
            }
        }

        private void gridGroupingControl1_SelectionChanged(object sender, EventArgs e)
        {
           
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {
        }

        private void add(string ID_KAR, string NM_KAR, string BAG, string DEPARTEMENT, string JK1, string BARCODE, string SECTION)
        {
            //SQL STMT
            SqlConnection conn = new SqlConnection(connString);
            const string sql = "INSERT INTO [BIODATA](NPK,NAMA_KARYAWAN,BAG,ID_DEPT,JENIS_KEL,BARCODE,SECTION) VALUES(@NPK12,@NM_KAR,@BAG2,@DEPT,@JK,@BARC,@SECT)";
            //cmd = new OdbcCommand(sql, conn);
            SqlCommand cmd = new SqlCommand(sql, conn);

            //ADD PARAMS
            cmd.Parameters.AddWithValue("@NPK12", ID_KAR);
            cmd.Parameters.AddWithValue("@NM_KAR", NM_KAR);
            cmd.Parameters.AddWithValue("@BAG2", BAG);
            cmd.Parameters.AddWithValue("@DEPT", DEPARTEMENT);
            cmd.Parameters.AddWithValue("@JK", JK1);
            cmd.Parameters.AddWithValue("@BARC", BARCODE);
            cmd.Parameters.AddWithValue("@SECT", SECTION);

            //OPEN CON AND EXEC INSERT
            try
            {
                conn.Open();
                if (cmd.ExecuteNonQuery() > 0)
                {
                    clearTxts();
                    MessageBox.Show(@"Successfully Inserted");
                }
                conn.Close();
                dt.Rows.Clear();
                retrieve();
                barcode();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                conn.Close();
            }
        }

        private void add1(string ID_KAR, string NM_KAR, string BAG, string DEPARTEMENT, string JK1, string BARCODE, string SECTION)
        {
            //SQL STMT
            SqlConnection conn = new SqlConnection(connString);
            const string sql = "INSERT INTO [BIODATA_KELUAR] (NPK,NAMA_KARYAWAN,BAG,ID_DEPT,JENIS_KEL,BARCODE,SECTION) VALUES(@NPK12,@NM_KAR,@BAG2,@DEPT,@JK,@BARC,@SECT)";
            //cmd = new OdbcCommand(sql, conn);
            SqlCommand cmd = new SqlCommand(sql, conn);

            //ADD PARAMS
            cmd.Parameters.AddWithValue("@NPK12", ID_KAR);
            cmd.Parameters.AddWithValue("@NM_KAR", NM_KAR);
            cmd.Parameters.AddWithValue("@BAG2", BAG);
            cmd.Parameters.AddWithValue("@DEPT", DEPARTEMENT);
            cmd.Parameters.AddWithValue("@JK", JK1);
            cmd.Parameters.AddWithValue("@BARC", BARCODE);
            cmd.Parameters.AddWithValue("@SECT", SECTION);

            //OPEN CON AND EXEC INSERT
            try
            {
                conn.Open();
                if (cmd.ExecuteNonQuery() > 0)
                {
                    clearTxts();
                    MessageBox.Show(@"Data Yang Dihapus Berhasil Dipindahkan");
                }
                conn.Close();
               
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                conn.Close();
            }
        }



        private void TAMBAHBTN_Click(object sender, EventArgs e)
        {
        }

        private void delete(string ID_KAR)
        {
            //SQL STATEMENTT
            SqlConnection conn = new SqlConnection(connString);
            string sql = "DELETE FROM [BIODATA] WHERE NPK='" + ID_KAR + "'";
            //cmd = new OdbcCommand(sql, conn);
            SqlCommand cmd = new SqlCommand(sql, conn);
            

            //ADD PARAMS
       //

            //'OPEN CONNECTION,EXECUTE DELETE,CLOSE CONNECTION
            try
            {
                conn.Open();
                //adapter = new OdbcDataAdapter(cmd);
               
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                adapter.DeleteCommand = conn.CreateCommand();
                adapter.DeleteCommand.CommandText = sql;
                
  

                //PROMPT FOR CONFIRMATION BEFORE DELETING
                if (MessageBox.Show(@"Are you sure to permanently delete this?", @"DELETE", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
                {
                    if (cmd.ExecuteNonQuery() > 0)
                    {
                        MessageBox.Show(@"Successfully deleted");
                        add1(NPK1.Text, NAMAK.Text, BAGI.Text, DEP.SelectedValue.ToString(), JK1.Text, BARC.Text, SECT.Text);

                    }
                }
                conn.Close();
                // gridGroupingControl1.DataSource = null;
                dt.Rows.Clear();
                retrieve();
                barcode();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                conn.Close();
            }
        }

        private void HAPUSBTN_Click(object sender, EventArgs e)
        {
        }

        private void update(string ID_KAR, string NM_KAR, string BAG, string DEPARTEMENT, string JK, string BARCODE, string SECTION)
        {
            //SQL STATEMENT
            SqlConnection conn = new SqlConnection(connString);
            string sql = "UPDATE [BIODATA]  SET NAMA_KARYAWAN='" + NM_KAR + "',BAG='" + BAG + "',ID_DEPT='" + DEPARTEMENT + "',JENIS_KEL='" + JK + "',BARCODE='" + BARCODE + "',SECTION='" + SECTION + "' WHERE NPK='" + ID_KAR + "'";
            //cmd = new OdbcCommand(sql, con);
            SqlCommand cmd = new SqlCommand(sql, conn);

            //OPEN CONNECTION,UPDATE,RETRIEVE DATAGRIDVIEW
            try
            {
                conn.Open();
                SqlDataAdapter adapter = new SqlDataAdapter(cmd)
                {
                    UpdateCommand = conn.CreateCommand()
                };
                adapter.UpdateCommand.CommandText = sql;
                if (adapter.UpdateCommand.ExecuteNonQuery() > 0)
                {
                    clearTxts();
                    MessageBox.Show(@"Successfully Updated");
                }
                conn.Close();

                //REFRESH DATA
                dt.Rows.Clear();
                retrieve();
                barcode();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                conn.Close();
            }
        }

        private void clearTxts()
        {
            NPK1.Text = "";
            NAMAK.Text = "";
            BAGI.Text = "";
            
            //BARC.Text = "";
        }

        public void barcode1()
        {
            BARC.Enabled = true;
        }

        private void EDITBTN_Click(object sender, EventArgs e)
        {
        }

        private void button1_Click(object sender, EventArgs e)
        {
        }

        private void NPK1_Enter(object sender, EventArgs e)
        {
        }

        private void dAILYToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DAILY dail = new DAILY();
            dail.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            add(NPK1.Text, NAMAK.Text, BAGI.Text, DEP.SelectedValue.ToString(), JK1.Text, BARC.Text, SECT.Text);
        }

        private void metroButton2_Click(object sender, EventArgs e)
        {
            update(NPK1.Text, NAMAK.Text, BAGI.Text, DEP.SelectedValue.ToString(), JK1.Text, BARC.Text, SECT.Text);
        }

        private void metroButton3_Click(object sender, EventArgs e)
        {
            delete(NPK1.Text);
            dt.DefaultView.RowFilter = null;
        }

        private void metroButton4_Click(object sender, EventArgs e)
        {
            clearTxts();
            dt.Rows.Clear();
            retrieve();
            departe();
            barcode();
            dt.DefaultView.RowFilter = null;
            if (BARC.Enabled == true)
            {
                BARC.Enabled = false;
            }
        }

        private void textBox1_KeyUp(object sender, KeyEventArgs e)
        {
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
        }

        private void rECORDToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RECORD RCD = new RECORD();
            RCD.ShowDialog();
        }

        private void BARC_Enter_1(object sender, EventArgs e)
        {
            if (BARC.Text == "")
            {
                barcode();
                BARC.Enabled = false;
            }
            else
            {
                BARC.Enabled = true;
            }
        }

        private void cALCULATIONDATEToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SETTING sett = new SETTING();
            sett.ShowDialog();
        }

        private void metroButton5_Click(object sender, EventArgs e)
        {
            if (metroComboBox1.SelectedItem.ToString() == "By Nama")
            {
                (gridGroupingControl1.DataSource as DataTable).DefaultView.RowFilter = string.Format("NAMA_KARYAWAN LIKE '{0}%'", metroTextBox1.Text);
            }
            else if (metroComboBox1.SelectedItem.ToString() == "By Departement")
            {
                (gridGroupingControl1.DataSource as DataTable).DefaultView.RowFilter = string.Format("DEPARTEMENT LIKE '{0}%'", metroTextBox1.Text);
            }
            else if (metroComboBox1.SelectedItem.ToString() == "By Bagian")
            {
                (gridGroupingControl1.DataSource as DataTable).DefaultView.RowFilter = string.Format("BAG LIKE '{0}%'", metroTextBox1.Text);
            }
            else
            {
                dt.DefaultView.RowFilter = string.Format("NPK LIKE '{0}%'", metroTextBox1.Text);
            }
        }

        private void metroComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            metroTextBox1.Enabled = true;
        }

        private void rETRIEVEToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            retrieve();
        }

        private void rEKAPBIODATAToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CEWCOW cewcow = new CEWCOW();
            cewcow.ShowDialog();
            
        }

        private void BARC_Click(object sender, EventArgs e)
        {
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {

        }
        
        private void dEToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DEPARTEMENT DPT = new DEPARTEMENT();
            DPT.ShowDialog();
        }

        private void DEP_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void DEP_Enter(object sender, EventArgs e)
        {
            
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void comboBox1_SelectionChangeCommitted(object sender, EventArgs e)
        {
           
        }

        private void metroButton6_Click(object sender, EventArgs e)
        {
            DEPARTEMENT DPT = new DEPARTEMENT();
            DPT.ShowDialog();
        }

        private void kARYAWANToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        #region Private Variables

        private GridGroupingExcelConverterControl excelConverter = new GridGroupingExcelConverterControl();
        private ExcelExportingOptions exportingOptions = new ExcelExportingOptions();

        #endregion Private Variables

        #region Event Customization

        private void excelConverter_QueryExportRowRange(object sender, QueryExportRowRangeEventArgs e)
        {
            GridTableDescriptor tableDescriptor = (GridTableDescriptor)e.Element.ParentTableDescriptor;
            int excelRowIndex = e.ExcelRange.Row;
            if (e.Element.Kind == DisplayElementKind.ColumnHeader)
            {
                for (int columnIndex = 0; columnIndex < tableDescriptor.VisibleColumns.Count; columnIndex++)
                {
                    IRange range = e.ExcelRange[excelRowIndex, e.ExcelRange.Column + columnIndex];
                    range.CellStyle.ColorIndex = Syncfusion.XlsIO.ExcelKnownColors.Rose;
                    range.CellStyle.Font.RGBColor = Color.DarkRed;
                    range.CellStyle.Font.FontName = "Segoe UI";
                    range.CellStyle.Font.Size = 10;
                    range.CellStyle.Font.Bold = true;
                    range.CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                }
            }
            else if (e.Element.Kind == DisplayElementKind.Caption)
            {
                IRange range = e.ExcelRange;
                range.CellStyle.ColorIndex = Syncfusion.XlsIO.ExcelKnownColors.Grey_40_percent;
                range.CellStyle.Font.RGBColor = Color.White;
                range.CellStyle.Font.FontName = "Segoe UI";
                range.CellStyle.Font.Size = 10;
                range.CellStyle.Font.Bold = true;
            }
            else if (e.Element.Kind == DisplayElementKind.Summary)
            {
                for (int columnIndex = 0; columnIndex < tableDescriptor.VisibleColumns.Count; columnIndex++)
                {
                    IRange range = e.ExcelRange[excelRowIndex, e.ExcelRange.Column + columnIndex];
                    range.CellStyle.ColorIndex = Syncfusion.XlsIO.ExcelKnownColors.Grey_25_percent;
                    range.CellStyle.Font.RGBColor = Color.White;
                    range.CellStyle.Font.FontName = "Segoe UI";
                    range.CellStyle.Font.Size = 10;
                    range.CellStyle.Font.Bold = true;
                }
            }
        }

        #endregion Event Customization

        #region Reading xml file

        private void ReadXml(DataSet ds, string xmlFileName)
        {
            for (int n = 0; n < 10; n++)
            {
                if (File.Exists(xmlFileName))
                {
                    ds.ReadXml(xmlFileName);
                    break;
                }
                xmlFileName = @"..\" + xmlFileName;
            }
        }


        #endregion Reading xml file

        private void eXCELToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DateTime date = DateTime.Today;

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel 97-2003 Workbook(*.Xls)|*.Xls|Excel Workbook(*.Xlsx)|*Xlsx";
            saveFileDialog.AddExtension = true;
            saveFileDialog.DefaultExt = ".Xlsx";
            saveFileDialog.FileName = "Rincian Karyawan " + date.ToString(string.Format("dd-MM-yyyy", date)) + " ";
            if (saveFileDialog.ShowDialog() == DialogResult.OK && saveFileDialog.CheckPathExists)
            {
                excelConverter.ExportToExcel(this.gridGroupingControl1, saveFileDialog.FileName, exportingOptions);

                if (MessageBox.Show("Do you wish to open the xls file now?", "Export to Excel", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    System.Diagnostics.Process proc = new System.Diagnostics.Process();
                    proc.StartInfo.FileName = saveFileDialog.FileName;
                    proc.Start();
                }
            }
        }

        private void gridGroupingControl1_SelectedRecordsChanged(object sender, SelectedRecordsChangedEventArgs e)
        {
            
        }

        private void gridGroupingControl1_TableControlCellClick(object sender, GridTableControlCellClickEventArgs e)
        {
            if (this.gridGroupingControl1.TableModel[e.Inner.RowIndex, e.Inner.ColIndex].CellType == "RowHeaderCell") //To check whether it is RowHeader

            {

                var s = this.gridGroupingControl1.Table.SelectedRecords;

                GridRangeInfoList s1 = this.gridGroupingControl1.TableModel.Selections.GetSelectedRows(true, true);//to get row index range

                foreach (GridRangeInfo info in s1)

                {

                    Element el = this.gridGroupingControl1.TableModel.GetDisplayElementAt(info.Top);

                    //to get cellvalue of particular column from selected row.  
                    string XNPK = el.GetRecord().GetValue("NPK").ToString();
                    string XNAMA = el.GetRecord().GetValue("NAMA_KARYAWAN").ToString();

                    string XBAG = el.GetRecord().GetValue("BAG").ToString();
                    string XDEPT = el.GetRecord().GetValue("DEPARTEMENT").ToString();

                    string XJK = el.GetRecord().GetValue("JENIS_KEL").ToString();
                    string XBARCODE = el.GetRecord().GetValue("BARCODE").ToString();

                    string XSECTION = el.GetRecord().GetValue("SECTION").ToString();

                    //TRANSFER TO TXT
                    NPK1.Text = XNPK;
                    NAMAK.Text = XNAMA;
                    BAGI.Text = XBAG;
                    DEP.Text = XDEPT + " - " + XSECTION;
                    JK1.Text = XJK;
                    BARC.Text = XBARCODE;
                    SECT.Text = XSECTION;
                }

            }
        }

        private void metroButton7_Click(object sender, EventArgs e)
        {

            this.Close();
            LOGIN LG = new LOGIN();
            LG.ShowDialog();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            label1.Text = DateTime.Now.ToLongTimeString();
        }

        private void aBSENSI2ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            COBA CB = new COBA();
            CB.ShowDialog();
        }

        private void metroButton8_Click(object sender, EventArgs e)
        {
            LOGIN LG = new LOGIN();
            LG.ShowDialog();
        }

        private void metroButton8_Click_1(object sender, EventArgs e)
        {
            BARC.Enabled = true;
            
        }

        private void fORM2ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form2 FRM2 = new Form2();
            FRM2.ShowDialog();
        }

        private void fORM3ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form3 FRM3 = new Form3();
            FRM3.ShowDialog();
        }

        private void fORM4ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form4 FRM4 = new Form4();
            FRM4.ShowDialog();
        }

        private void kARYAWANKELUARToolStripMenuItem_Click(object sender, EventArgs e)
        {
            KARYKELUAR KARYKEL = new KARYKELUAR();
            KARYKEL.ShowDialog();
        }
    }
}