using System;
using MetroFramework.Forms;
using MetroFramework.Fonts;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Data.Sql;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace JPT_KARPRO
{
    public partial class LOGIN : MetroFramework.Forms.MetroForm
    {
        public LOGIN()
        {
            InitializeComponent();
        }

        private void LOGIN_Load(object sender, EventArgs e)
        {
            textBox1.Focus();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" || textBox2.Text == "")
            {
                MessageBox.Show("username atau password harap diisi !");
                textBox1.Focus();
            }
            else
            {
                string connectionString = "Data Source=ANNNIIXX;Initial Catalog=JAPER-DATA;integrated security = true";

                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("SELECT username,pass from LOGIN where username='" + textBox1.Text + "' AND pass='" + textBox2.Text + "'", con))
                    {
                        SqlDataAdapter DA = new SqlDataAdapter(cmd);

                        DataTable dt = new DataTable();
                        DA.Fill(dt);

                        if (dt.Rows.Count > 0)
                        {
                            foreach (DataRow dr in dt.Rows)
                            {
                                if (dr["username"].ToString() == "admin")
                                {
                                    MessageBox.Show("Login Success! Welcome Administrator !");
                                    SplashForm SF = new SplashForm();
                                    SF.ShowDialog();
                                    this.Hide();
                                    con.Close();
                                }
                                else if (dr["username"].ToString() == "nia")  
                                {
                                    MessageBox.Show("Login Success! Welcome MBAK NIA !");
                                    Form6 FBK = new Form6();
                                    FBK.ShowDialog();
                                    this.Hide();
                                    con.Close();
                                   
                                }
                                else if (dr["username"].ToString() == "nining")
                                {
                                    MessageBox.Show("Login Success! Welcome MBAK NINING !");
                                    IZIN ijin = new IZIN();
                                    ijin.ShowDialog();

                                    this.Hide();
                                    con.Close();

                                }




                            }
                        }
                        else
                        {
                            MessageBox.Show("Username atau password salah !");
                            textBox1.Text = "";
                            textBox2.Text = "";

                        }
                    }
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
