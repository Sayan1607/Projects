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
    public partial class EditInputData : Form
    {
        SqlConnection sqlcon = new SqlConnection(@"Data Source=localhost;Initial Catalog=ProjectHealth;Integrated Security=True;");
        public EditInputData()
        {
            InitializeComponent();
        }

        private void EditInputData_Load(object sender, EventArgs e)
        {
            try
            {
                sqlcon.Open();
                SqlDataAdapter sda = new SqlDataAdapter("SELECT * FROM [dbo].[ProjectHelathEntry]", sqlcon);
                DataTable dt = new DataTable();
                sda.Fill(dt);
                dgv4.DataSource = dt;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
                sqlcon.Close();
            }
            finally
            {
                sqlcon.Close();
            }
        }

       
        private void lblUD2_Click(object sender, EventArgs e)
        {

        }

        private void txtUD2_TextChanged(object sender, EventArgs e)
        {

        }

        private void lblUD3_Click(object sender, EventArgs e)
        {

        }

        private void txtUD3_TextChanged(object sender, EventArgs e)
        {

        }

        private void lbl330Ind_Click(object sender, EventArgs e)
        {

        }

        private void txt330Ind_TextChanged(object sender, EventArgs e)
        {

        }

        private void lblGapLimit_Click(object sender, EventArgs e)
        {

        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            try
            {
                sqlcon.Open();
                //String var = "INSERT INTO [dbo].[ProjectHelathEntry] values (" + Convert.ToInt32(txtItemNo.Text) + "," + Convert.ToDecimal(txtPrclInd.Text) + "," + Convert.ToDecimal(txtUD1.Text) + "," + Convert.ToDecimal(txtUD2.Text) + "," + Convert.ToDecimal(txtUD3.Text) + "," + Convert.ToDecimal(txt330Ind.Text) + "," + Convert.ToInt32(txtGapLimit.Text) + "," + Convert.ToDecimal(txtIRES.Text) + "," + Convert.ToDecimal(txtS1C2.Text) + "," + Convert.ToDecimal(txtPRCLNXTDCLDIF.Text) + ")";
                String var = "UPDATE [dbo].[ProjectHelathEntry] SET [PRCL_IND] =" + Convert.ToDecimal(txtPrclInd.Text) + "," + " [UD1] =" + Convert.ToDecimal(txtUD1.Text) + "," +" [UD2] =" + Convert.ToDecimal(txtUD2.Text) + "," + " [UD3] =" + Convert.ToDecimal(txtUD3.Text) + "," + " [330_IND] =" + Convert.ToDecimal(txt330Ind.Text) + "," + " [GAP_LIMIT] =" + Convert.ToDecimal(txtGapLimit.Text) + "," + " [IRES] =" + Convert.ToDecimal(txtIRESNew.Text) + "," + " [S1C2] =" + Convert.ToDecimal(txtS1C2.Text) + "," + " [PRCL-NXTDCL-DIF] =" + Convert.ToDecimal(txtPRCLNXTDCLDIF.Text) + " WHERE ITEM_NO ='" + Convert.ToInt32(txtItemNo.Text) + "'";
                //MessageBox.Show(var);
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = sqlcon;
                cmd.CommandText = var;
                cmd.ExecuteNonQuery();
                //SqlDataAdapter sda = new SqlDataAdapter("INSERT INTO [dbo].[ProjectHelathEntry] values (" + Convert.ToInt32(txtItemNo.Text) + "," + Convert.ToDecimal(txtPrclInd.Text) + "," + Convert.ToDecimal(txtUD1.Text)+"," + Convert.ToDecimal(txtUD2.Text)+"," + Convert.ToDecimal(txtUD3.Text) + "," + Convert.ToDecimal(txt330Ind.Text) + "," + Convert.ToInt32(txtGapLimit.Text)+")", sqlcon);

                //DataTable dt = new DataTable();
                //sda.Fill(dt);
                //dgv1.DataSource = dt;
                MessageBox.Show("One Record Updated Successfully.");
                this.Close();
                Main ms = new Main();
                ms.Show();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
                sqlcon.Close();
            }
            finally
            {
                sqlcon.Close();
            }
        }

        private void btnSelect_Click(object sender, EventArgs e)
        {
            try
            {
                sqlcon.Open();
                SqlDataAdapter sda = new SqlDataAdapter("SELECT * FROM [dbo].[ProjectHelathEntry] WHERE ITEM_NO ='" + Convert.ToInt32(txtItemNo.Text) + "'", sqlcon);
                DataTable dt = new DataTable();
                sda.Fill(dt);
                dgv4.DataSource = dt;
                //MessageBox.Show(dt.Rows.Count.ToString());

                for (int index = 0; index < dt.Rows.Count; index++)
                {
                    txtPrclInd.Text = dt.Rows[0]["PRCL_IND"].ToString();
                    txtUD1.Text = dt.Rows[0]["UD1"].ToString();
                    txtUD2.Text = dt.Rows[0]["UD2"].ToString();
                    txtUD3.Text = dt.Rows[0]["UD3"].ToString();
                    txt330Ind.Text = dt.Rows[0]["330_IND"].ToString();
                    txtGapLimit.Text = dt.Rows[0]["GAP_LIMIT"].ToString();
                    txtIRESNew.Text = dt.Rows[0]["IRES"].ToString();
                    txtS1C2.Text = dt.Rows[0]["S1C2"].ToString();
                    txtPRCLNXTDCLDIF.Text = dt.Rows[0]["PRCL-NXTDCL-DIF"].ToString();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
                sqlcon.Close();
            }
            finally
            {
                sqlcon.Close();
            }

        }
    }
}
