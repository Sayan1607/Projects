using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;

namespace ProjectHealthApplication
{
    public partial class LogOn : Form
    {
        //static string str = ConfigurationManager.ConnectionStrings["SQLConnstring"].ConnectionString;
        //SqlConnection sqlconn = new SqlConnection(str);
        SqlConnection sqlconn = new SqlConnection(@"Data Source=L99816120\MSSQLSERVER01;Initial Catalog=ProjectHealth;Integrated Security=True;");
        public static string UsrID = "";
        public LogOn()
        {
            InitializeComponent();
        }

    private void LogIn_Click(object sender, EventArgs e)
    {
        try
        {

            sqlconn.Open();
            SqlDataAdapter sda = new SqlDataAdapter("select count(*) from dbo.LogonTable where Id ='" + textBox1.Text + "' and Password ='" + textBox2.Text + "'", sqlconn);
            DataTable dtr = new DataTable();
            sda.Fill(dtr);
            if (dtr.Rows[0][0].ToString() == "1")
            {
                sqlconn.Close();
                this.Hide();
                Main mm = new Main();
                mm.Show();
            }
            else
            {
                MessageBox.Show("Please check your UserName and Password");
                textBox1.Clear();
                textBox2.Clear();
                sqlconn.Close();
            }
        }
        catch (Exception ex)
        {
            sqlconn.Close();
            MessageBox.Show(ex.Message.ToString());
        }
    }

    private void Exit_Click(object sender, EventArgs e)
    {
        this.Close();
        System.Windows.Forms.Application.Exit();
    }

    private void btnChngPasswd_Click(object sender, EventArgs e)
    {
        UsrID = textBox1.Text.ToString();
        this.Hide();
        ChangePassword cs = new ChangePassword();
        cs.Show();

    }

    }
}
