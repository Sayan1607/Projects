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

namespace ProjectHealthApplication
{
    public partial class ChangePassword : Form
    {
        SqlConnection sqcon = new SqlConnection(@"Data Source=localhost;Initial Catalog=ProjectHealth;Integrated Security=True;");
        //string UsrID = "";
        public ChangePassword()
        {
            InitializeComponent();
        }

        private void btnSubmit_Click(object sender, EventArgs e)
        {
            try
            {
                sqcon.Open();
                if(txtChngdPaswwd.Text.ToString()=="" || txtRptChngdPasswd.Text.ToString()=="")
                {
                    MessageBox.Show("Password cannot be blank. Please Reenter the password");
                    txtChngdPaswwd.Clear();
                    txtRptChngdPasswd.Clear();
                    txtOldPasswd.Clear();
                }
                else if(txtChngdPaswwd.Text.ToString() == txtRptChngdPasswd.Text.ToString())
                {
                    String var = "Update dbo.LogonTable set Password ='" + txtChngdPaswwd.Text + "' where Id ='" + LogOn.UsrID + "'";
                    //MessageBox.Show(var);
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = sqcon;
                    cmd.CommandText = var;
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Password changed successfully.");
                    this.Hide();
                    LogOn lo = new LogOn();
                    lo.Show();
                }
                else
                {
                    MessageBox.Show("Please reenter same password.");
                    txtChngdPaswwd.Clear();
                    txtRptChngdPasswd.Clear();
                    txtOldPasswd.Clear();
                }
               
            }
            catch (Exception ex)
            {
                sqcon.Close();
                MessageBox.Show(ex.Message.ToString());
            }
            finally
            {
                sqcon.Close();
            }

        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void ChangePassword_Load(object sender, EventArgs e)
        {
            //UsrID = LogOn.UsrID;
        }
    }
}
