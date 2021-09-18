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
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Configuration;
using System.Data.OleDb;
using System.Globalization;

namespace ProjectHealthApplication
{
    public partial class Main : Form
    {
        //constr = ConfigurationManager.ConnectionStrings["getconn"].ToString();
        //sqlcon = new SqlConnection(constr);
        ConnectionStringSettingsCollection settings = ConfigurationManager.ConnectionStrings;
        SqlConnection sqlcon = new SqlConnection(@"Data Source=localhost;Initial Catalog=ProjectHealth;Integrated Security=True;");
        public Main()
        {
            InitializeComponent();
            try
            {
                sqlcon.Open();
                SqlDataAdapter sda = new SqlDataAdapter("SELECT * FROM [dbo].[ProjectHelathEntryResult]", sqlcon);
                DataTable dt = new DataTable();
                sda.Fill(dt);
                dgv3.DataSource = dt;
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

        private void btnDeatilsEntry_Click(object sender, EventArgs e)
        {
            try
            {
                sqlcon.Open();
                if (String.IsNullOrEmpty(txtUD1.Text))
                {
                    txtUD1.Text = Convert.ToString(0.03);
                }
                if (String.IsNullOrEmpty(txtUD2.Text))
                {
                    txtUD2.Text = Convert.ToString(0.05);
                }
                
                if (String.IsNullOrEmpty(txtUD3.Text))
                {
                    txtUD3.Text = Convert.ToString(0.10);
                }
                
                if (String.IsNullOrEmpty(txtGapLimit.Text))
                {
                    txtGapLimit.Text = Convert.ToString(1.00);
                }
                if (String.IsNullOrEmpty(txtPRFININD.Text))
                {
                    txtPRFININD.Text = Convert.ToString(1.00);
                }

                if (String.IsNullOrEmpty(txtUD4.Text))
                {
                    txtUD4.Text = Convert.ToString(0.03);
                }
                if (String.IsNullOrEmpty(txtUD5.Text))
                {
                    txtUD5.Text = Convert.ToString(0.05);
                }

                if (String.IsNullOrEmpty(txtUD6.Text))
                {
                    txtUD6.Text = Convert.ToString(0.10);
                }
                if (String.IsNullOrEmpty(txtNoofCells.Text))
                {
                    txtNoofCells.Text = Convert.ToString(1);
                }

                String vars = "SP_DeletePP_ProjectEntry";
                    //MessageBox.Show(var);
                    SqlCommand cmds = new SqlCommand();
                    cmds.Connection = sqlcon;
                    cmds.CommandText = vars;
                    cmds.ExecuteNonQuery();

                //MessageBox.Show(dtP.Value.ToString("dd/MM/yyyy"));
                String var = "INSERT INTO [dbo].[ProjectHelathEntry] values (" + Convert.ToInt32(txtItemNo.Text) + "," + Convert.ToDecimal(txtPrclInd.Text) + "," + Convert.ToDecimal(txtUD1.Text) + "," + Convert.ToDecimal(txtUD2.Text) + "," + Convert.ToDecimal(txtUD3.Text)+ "," + Convert.ToDecimal(txtUD4.Text) + "," + Convert.ToDecimal(txtUD5.Text) + "," + Convert.ToDecimal(txtUD6.Text) + "," + Convert.ToDecimal(txt330Ind.Text) + "," + Convert.ToDecimal(txtGapLimit.Text) + "," + Convert.ToDecimal(txtIRESNew.Text) + "," + Convert.ToDecimal(txtPRS1C1.Text) + "," + Convert.ToDecimal(txtS1C2.Text) + "," + Convert.ToDecimal(txtPRS1C3.Text) +"," + Convert.ToDecimal(txtPRCLNXTDCLDIF.Text)+","  +"'" + dtP.Value.ToString("yyyy-MM-dd").Trim() +"'" +"," + Convert.ToDecimal(txtPRFININD.Text) +"," +  Convert.ToInt32(txtNoofCells.Text) +  ")";
                //MessageBox.Show(var.ToString());
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = sqlcon;
                cmd.CommandText = var;
                cmd.ExecuteNonQuery();
                //SqlDataAdapter sda = new SqlDataAdapter("INSERT INTO [dbo].[ProjectHelathEntry] values (" + Convert.ToInt32(txtItemNo.Text) + "," + Convert.ToDecimal(txtPrclInd.Text) + "," + Convert.ToDecimal(txtUD1.Text)+"," + Convert.ToDecimal(txtUD2.Text)+"," + Convert.ToDecimal(txtUD3.Text) + "," + Convert.ToDecimal(txt330Ind.Text) + "," + Convert.ToInt32(txtGapLimit.Text)+")", sqlcon);

                //DataTable dt = new DataTable();
                //sda.Fill(dt);
                //dgv1.DataSource = dt;
                MessageBox.Show("One Record Inserted Successfully.");
                dgv3.DataSource = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
                sqlcon.Close();
            }
            finally
            {
                sqlcon.Close();
                txtItemNo.Text = "";
                txtPrclInd.Text = "";
                txtUD1.Text = "";
                txtUD2.Text = "";
                txtUD3.Text = "";
                txtUD4.Text = "";
                txtUD5.Text = "";
                txtUD6.Text = "";
                txt330Ind.Text = "";
                txtGapLimit.Text = "";
                txtIRESNew.Text = "";
                txtS1C2.Text = "";
                txtPRCLNXTDCLDIF.Text = "";
                txtPRS1C1.Text = "";
                txtPRS1C3.Text = "";
            }
            
        }

        private void btnShowData_Click(object sender, EventArgs e)
        {
            try
            {
                sqlcon.Open();
                SqlDataAdapter sda = new SqlDataAdapter("SELECT * FROM [dbo].[ProjectHelathEntry]", sqlcon);
                DataTable dt = new DataTable();
                sda.Fill(dt);
                dgv1.DataSource = dt;

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

        private void btnShowPPFile_Click(object sender, EventArgs e)
        {
            try
            {
                sqlcon.Open();
                SqlDataAdapter sda = new SqlDataAdapter("SELECT * FROM [dbo].[PP-FILE]", sqlcon);
                DataTable dt = new DataTable();
                sda.Fill(dt);
                dgv2.DataSource = dt;

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

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
            System.Windows.Forms.Application.Exit();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
        public  void ExportRows(DataGridView sender, string fileName)
        {
            if (sender.RowCount > 0)
            {
                var sb = new StringBuilder();

                var headers = sender.Columns.Cast<DataGridViewColumn>();
                sb.AppendLine(string.Join("\t", headers.Select(column => column.HeaderText)));

                foreach (DataGridViewRow row in sender.Rows)
                {
                    if (!row.IsNewRow == true)
                    {
                        var cells = row.Cells.Cast<DataGridViewCell>();
                        sb.AppendLine(string.Join("\t", cells.Select(cell => cell.Value)));
                    }
                }
                System.IO.File.WriteAllText(fileName, sb.ToString());
            }
            else
            {
                MessageBox.Show("Export Issue Occured.");
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCalculate_Click(object sender, EventArgs e)
        {
            try
            {
                sqlcon.Open();
                String var = "SP_LoadFinalHealthData";
                //MessageBox.Show(var);
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = sqlcon;
                cmd.CommandText = var;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@Minnoofcells", SqlDbType.VarChar).Value = txtMinNoOfCelssP1.Text;
                cmd.CommandTimeout = 3600000;
                cmd.ExecuteNonQuery();
                MessageBox.Show("Calculation Completed Successfully.");
                SqlDataAdapter sda = new SqlDataAdapter("SELECT [ITEM_NO],	[PR-DATE],	[PRCL-IND],	[330-IND],	[S1C2],	[UD1],	[PRCL/PRCL-IND-CNT1],	[UD2],	[PRCL/PRCL-IND-CNT2],	[UD3],	[PRCL/PRCL-IND-CNT3],	[PRCL/C2-CNT1],	[PRCL/C2-CNT2],	[PRCL/C3-CNT3],	[RED/INC],	[PP-RNG-UD],	[GAP-LMT],	[IND-GAP],	[LIND],	[MIND],	[Res9],	[Res8],	[Res7],	[Res6],	[Res5],	[Res4],	[Res3],	[Res2],	[Res1],	[TOT], [ALL-CNT(P/N)],	[GREEN],[ACC/TOT],[3COL/NET], [ACC/3COL],	[COMM],	[IRES], [PRFIN-IND],[UD1] AS [_UD1],[PRCL/330IND-R1-C],[UD2] AS [_UD2],[PRCL/330IND-R2-C],[UD3] AS [_UD3],[PRCL/330IND-R3-C],[PR-S1C1],[PR-S1C2],[PR-S1C3],[MinNoOfCells],[PR-DATE],[2COL/NET],	[4COL/NET],	[3COL/ST/ALIGNED],	[3COL/ST/GAP],	[S1/LH/CNT/SUM],	[S2/LH/CNT/SUM],	[S3/LH/CNT/SUM],	[LH-MINUS-SUM/S1+S2+S3(ADD ONLY MINUSES, ELSE 0)], [NO-OF-Ls/NO-OF-Hs],[ACC/X],[RED/AVG-POS],[RED/AVG-NEG],NULL,[RED-P/13&12&23ROWS-TR],[RED-N/13&12&23ROWS-TR],NULL,[UP/SUM],[DW/SUM],NULL,[UP/CNT],[DW/CNT],NULL,[13&12&23/UP/CNT],[13&12&23/ DW/CNT],NULL,[UP/AMT-AVG],[DW/AMT-AVG],[ACC],NULL,[PP-PRCL-PPOP/33IND-PPCL/NET-AVG] FROM[dbo].[ProjectHelathEntryResult] ORDER BY ITEM_NO ASC, [RED/INC] desc", sqlcon);
                DataSet ds = new DataSet();
                sda.Fill(ds, "ProjectHelathEntryResult");
                dgv3.DataSource = ds;
                dgv3.DataMember = "ProjectHelathEntryResult";

                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                using (ExcelPackage excelPackage = new ExcelPackage())
                {
                    
                    //Set some properties of the Excel document
                    excelPackage.Workbook.Properties.Author = "Sayan";
                    excelPackage.Workbook.Properties.Title = "P1-INDIVIDUAL-DETAIL-OUTPUTFILE-V1";
                    excelPackage.Workbook.Properties.Subject = "P1-INDIVIDUAL-DETAIL-OUTPUTFILE-V1";
                    excelPackage.Workbook.Properties.Created = DateTime.Now;

                    //Create the WorkSheet
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("P1-DETAIL-OUTPUT-FILE");
                   
                    for (int i = 0; i <= dgv3.Columns.Count - 1; i++)
                    {
                        worksheet.Cells[1, i + 1].Value = dgv3.Columns[i].HeaderText;
                    }

                    /*And the information of your data*/
                    for (int i = 0; i <= dgv3.RowCount - 1; i++)
                    {
                        for (int j = 0; j <= dgv3.ColumnCount - 1; j++)
                        {
                            DataGridViewCell cell = dgv3[j, i];
                            worksheet.Cells[i + 2, j + 1].Value = cell.Value;
                            
                        }
                    }

                    //Save your file
                    FileInfo fi = new FileInfo(@"C:\ProjectHealthApplication\Export\P1-INDIVIDUAL-DETAIL-OUTPUTFILE-V1.xlsx");
                    excelPackage.SaveAs(fi);
                }
                try
                {
                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                    ExcelPackage ExcelPkg = new ExcelPackage();

                    FileInfo fi = new FileInfo(@"C:\ProjectHealthApplication\Export\P1-INDIVIDUAL-DETAIL-OUTPUTFILE-V1.xlsx");
                    using (ExcelPackage excelPackage = new ExcelPackage(fi))
                    {
                        //Get a WorkSheet by name. If the worksheet doesn't exist, throw an exeption
                        ExcelWorksheet namedWorksheet = excelPackage.Workbook.Worksheets["P1-DETAIL-OUTPUT-FILE"];

                        //If you don't know if a worksheet exists, you could use LINQ,
                        //So it doesn't throw an exception, but return null in case it doesn't find it
                        ExcelWorksheet anotherWorksheet =
                            excelPackage.Workbook.Worksheets.FirstOrDefault(x => x.Name == "P1-DETAIL-OUTPUT-FILE");


                        //ExcelRange Rng1= namedWorksheet.Cells[2,1, dgv3.RowCount,2]
                        //int j = 0;
                        for (int k = 0; k < dgv3.ColumnCount; k++)
                        {

                            using (ExcelRange Rng = namedWorksheet.Cells[1, 1, 1, k + 1])
                            {
                                namedWorksheet.Row(1).Height = 40;
                                namedWorksheet.Column(k + 1).Width = 6;
                                namedWorksheet.Column(k + 1).Style.WrapText = true;
                                namedWorksheet.Cells[1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                                //namedWorksheet.Cells[1, 2].Style.
                                namedWorksheet.Cells[1, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[1, 2].Style.Fill.BackgroundColor.SetColor(Color.Yellow);

                                namedWorksheet.Cells[1, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[1, 3].Style.Fill.BackgroundColor.SetColor(Color.Gold);

                                namedWorksheet.Cells[1, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[1, 4].Style.Fill.BackgroundColor.SetColor(Color.Yellow);

                                //namedWorksheet.Cells[1, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                //namedWorksheet.Cells[1, 6].Style.Fill.BackgroundColor.SetColor(Color.DeepSkyBlue);

                                namedWorksheet.Cells[1, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[1, 6].Style.Fill.BackgroundColor.SetColor(Color.DeepSkyBlue);

                                namedWorksheet.Cells[1, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[1, 7].Style.Fill.BackgroundColor.SetColor(Color.LightCyan);

                                namedWorksheet.Cells[1, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[1, 8].Style.Fill.BackgroundColor.SetColor(Color.DeepSkyBlue);

                                namedWorksheet.Cells[1, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[1, 9].Style.Fill.BackgroundColor.SetColor(Color.LightCyan);

                                namedWorksheet.Cells[1, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[1, 10].Style.Fill.BackgroundColor.SetColor(Color.DeepSkyBlue);

                                namedWorksheet.Cells[1, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[1, 11].Style.Fill.BackgroundColor.SetColor(Color.LightCyan);

                                namedWorksheet.Cells[1, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[1, 12].Style.Fill.BackgroundColor.SetColor(Color.DarkGray);
                                namedWorksheet.Cells[1, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[1, 13].Style.Fill.BackgroundColor.SetColor(Color.DarkGray);
                                namedWorksheet.Cells[1, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[1, 14].Style.Fill.BackgroundColor.SetColor(Color.DarkGray);
                                

                                namedWorksheet.Cells[1, 15].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[1, 15].Style.Fill.BackgroundColor.SetColor(Color.Violet);

                                namedWorksheet.Cells[1, 16].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[1, 16].Style.Fill.BackgroundColor.SetColor(Color.PeachPuff);

                                namedWorksheet.Cells[1, 19].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[1, 19].Style.Fill.BackgroundColor.SetColor(Color.LightSkyBlue);
                                namedWorksheet.Cells[1, 20].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[1, 20].Style.Fill.BackgroundColor.SetColor(Color.LightSkyBlue);

                                namedWorksheet.Cells[1, 30].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[1, 30].Style.Fill.BackgroundColor.SetColor(Color.Green);
                                namedWorksheet.Cells[1, 31].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[1, 31].Style.Fill.BackgroundColor.SetColor(Color.Green);
                                //namedWorksheet.Cells[1, 32].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                //namedWorksheet.Cells[1, 32].Style.Fill.BackgroundColor.SetColor(Color.Green);

                                namedWorksheet.Cells[1, 34].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[1, 34].Style.Fill.BackgroundColor.SetColor(Color.Green);
                                namedWorksheet.Cells[1, 37].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[1, 37].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                                namedWorksheet.Cells[1, 39].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[1, 39].Style.Fill.BackgroundColor.SetColor(Color.LightGoldenrodYellow);
                                namedWorksheet.Cells[1, 40].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[1, 40].Style.Fill.BackgroundColor.SetColor(Color.LightGoldenrodYellow);
                                namedWorksheet.Cells[1, 41].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[1, 41].Style.Fill.BackgroundColor.SetColor(Color.LightGoldenrodYellow);
                                namedWorksheet.Cells[1, 42].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[1, 42].Style.Fill.BackgroundColor.SetColor(Color.LightGoldenrodYellow);
                                namedWorksheet.Cells[1, 43].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[1, 43].Style.Fill.BackgroundColor.SetColor(Color.LightGoldenrodYellow);
                                namedWorksheet.Cells[1, 44].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[1, 44].Style.Fill.BackgroundColor.SetColor(Color.LightGoldenrodYellow);

                                namedWorksheet.Cells[1, 62].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[1, 62].Style.Fill.BackgroundColor.SetColor(Color.Blue);
                                namedWorksheet.Cells[1, 65].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[1, 65].Style.Fill.BackgroundColor.SetColor(Color.Blue);
                                namedWorksheet.Cells[1, 68].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[1, 68].Style.Fill.BackgroundColor.SetColor(Color.Blue);
                                namedWorksheet.Cells[1, 71].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[1, 71].Style.Fill.BackgroundColor.SetColor(Color.Blue);
                                namedWorksheet.Cells[1, 74].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[1, 74].Style.Fill.BackgroundColor.SetColor(Color.Green);
                                namedWorksheet.Cells[1, 78].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[1, 78].Style.Fill.BackgroundColor.SetColor(Color.Blue);
                                

                                namedWorksheet.View.FreezePanes(2, 1);
                                //Rng.AutoFitColumns();

                                Rng.Style.Font.Size = 8;
                                Rng.Style.Font.Bold = true;
                                Rng.Style.Font.Color.SetColor(Color.Red);
                                Rng.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                Rng.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                Rng.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                Rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                Rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            }
                            
                            using (ExcelRange Rng = namedWorksheet.Cells[2, (k+1), dgv3.RowCount, (k + 1)])
                            {
                               
                                if(k==2)
                                {
                                    Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    Rng.Style.Fill.BackgroundColor.SetColor(Color.Gold);
                                   
                                }
                                else if(k==3)
                                {
                                    Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    Rng.Style.Fill.BackgroundColor.SetColor(Color.Green);
                                }
                                else if (k == 5 || k == 7 || k == 9)
                                {
                                    Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    Rng.Style.Fill.BackgroundColor.SetColor(Color.DeepSkyBlue);
                                }
                                else if (k == 6 || k == 8 || k == 10)
                                {
                                    Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    Rng.Style.Fill.BackgroundColor.SetColor(Color.LightCyan);
                                }
                                else if (k == 11 || k == 12 || k == 13)
                                {
                                    Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    Rng.Style.Fill.BackgroundColor.SetColor(Color.DarkGray);
                                }
                                else if (k == 14)
                                {
                                    Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    Rng.Style.Fill.BackgroundColor.SetColor(Color.Violet);
                                    
                                }
                                else if (k == 15)
                                {
                                    Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    Rng.Style.Fill.BackgroundColor.SetColor(Color.PeachPuff);
                                }
                                else if (k == 18 || k == 19)
                                {
                                    Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    Rng.Style.Fill.BackgroundColor.SetColor(Color.LightSkyBlue);
                                }
                                else if (k >=20 && k<=28)
                                {
                                    Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    Rng.Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                                       
                                }
                                else if (k >= 29 && k <= 30)
                                {
                                    Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    Rng.Style.Fill.BackgroundColor.SetColor(Color.Green);
                                }
                                else if (k == 33)
                                {
                                    Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    Rng.Style.Fill.BackgroundColor.SetColor(Color.Green);
                                }
                                else if (k == 36)
                                {
                                    Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    Rng.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                                }
                                else if (k == 38 || k == 40 || k == 42)
                                {
                                    Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    Rng.Style.Fill.BackgroundColor.SetColor(Color.LawnGreen);
                                }
                                else if (k == 61 || k == 64 || k == 67 || k == 70 || k == 77)
                                {
                                    Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    Rng.Style.Fill.BackgroundColor.SetColor(Color.Blue);
                                }
                                else if (k == 73)
                                {
                                    Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    Rng.Style.Fill.BackgroundColor.SetColor(Color.Green);
                                }
                                else
                                {
                                    Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    Rng.Style.Fill.BackgroundColor.SetColor(Color.White);
                                }

                                if(k == 61 || k == 64 || k == 67 || k == 70 || k == 73 || k == 77)
                                {
                                    namedWorksheet.Column(k + 1).Width = 2;
                                }
                                else
                                {
                                    namedWorksheet.Column(k + 1).Width = 6;
                                }
                                
                                namedWorksheet.Column(k + 1).Style.WrapText = true;
                                Rng.Style.Font.Size = 8;
                                Rng.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                Rng.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                Rng.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                Rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                Rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                //Rng.AutoFitColumns();
                               
                            }

                            
                        }
                        for (int i= 2;i<= dgv3.RowCount;i++)
                        {
                            if(namedWorksheet.Cells[i, 15].Value.ToString() =="RED")
                            {
                                if(Convert.ToDecimal(namedWorksheet.Cells[i, 4].Value.ToString()) < Convert.ToDecimal(namedWorksheet.Cells[i, 3].Value.ToString()))
                                {
                                    namedWorksheet.Cells[i, 15].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[i, 15].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                    namedWorksheet.Cells[i, 16].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[i, 16].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                    namedWorksheet.Cells[i, 17].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[i, 17].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                    namedWorksheet.Cells[i, 18].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[i, 18].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                    //namedWorksheet.Cells[i, 19].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    //namedWorksheet.Cells[i, 19].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                    //namedWorksheet.Cells[i, 20].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    //namedWorksheet.Cells[i, 20].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                    namedWorksheet.Cells[i, 21].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[i, 21].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                    namedWorksheet.Cells[i, 22].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[i, 22].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                    namedWorksheet.Cells[i, 23].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[i, 23].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                    namedWorksheet.Cells[i, 24].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[i, 24].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                    namedWorksheet.Cells[i, 25].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[i, 25].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                    namedWorksheet.Cells[i, 26].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[i, 26].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                    namedWorksheet.Cells[i, 27].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[i, 27].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                    namedWorksheet.Cells[i, 28].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[i, 28].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                    namedWorksheet.Cells[i, 29].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[i, 29].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                    //namedWorksheet.Cells[i, 30].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    //namedWorksheet.Cells[i, 30].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                    //namedWorksheet.Cells[i, 31].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    //namedWorksheet.Cells[i, 31].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                    namedWorksheet.Cells[i, 32].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[i, 32].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                    namedWorksheet.Cells[i, 33].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[i, 33].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                    //namedWorksheet.Cells[i, 34].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    //namedWorksheet.Cells[i, 34].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                    namedWorksheet.Cells[i, 35].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[i, 35].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);

                                    namedWorksheet.Cells[i, 62].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[i, 62].Style.Fill.BackgroundColor.SetColor(Color.Blue);
                                    namedWorksheet.Cells[i, 65].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[i, 65].Style.Fill.BackgroundColor.SetColor(Color.Blue);
                                    namedWorksheet.Cells[i, 68].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[i, 68].Style.Fill.BackgroundColor.SetColor(Color.Blue);
                                    namedWorksheet.Cells[i, 71].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[i, 71].Style.Fill.BackgroundColor.SetColor(Color.Blue);
                                    namedWorksheet.Cells[i, 74].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[i, 74].Style.Fill.BackgroundColor.SetColor(Color.Green);
                                    namedWorksheet.Cells[i, 78].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[i, 78].Style.Fill.BackgroundColor.SetColor(Color.Blue);
                                    

                                }
                            }
                            if (namedWorksheet.Cells[i, 15].Value.ToString() == "INC")
                            {
                                if (Convert.ToDecimal(namedWorksheet.Cells[i, 4].Value.ToString()) > Convert.ToDecimal(namedWorksheet.Cells[i, 3].Value.ToString()))
                                {
                                    //MessageBox.Show("INC Hi...");
                                    namedWorksheet.Cells[i, 15].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[i, 15].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                    namedWorksheet.Cells[i, 16].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[i, 16].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                    namedWorksheet.Cells[i, 17].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[i, 17].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                    namedWorksheet.Cells[i, 18].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[i, 18].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                    //namedWorksheet.Cells[i, 19].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    //namedWorksheet.Cells[i, 19].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                    //namedWorksheet.Cells[i, 20].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    //namedWorksheet.Cells[i, 20].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                    namedWorksheet.Cells[i, 21].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[i, 21].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                    namedWorksheet.Cells[i, 22].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[i, 22].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                    namedWorksheet.Cells[i, 23].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[i, 23].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                    namedWorksheet.Cells[i, 24].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[i, 24].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                    namedWorksheet.Cells[i, 25].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[i, 25].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                    namedWorksheet.Cells[i, 26].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[i, 26].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                    namedWorksheet.Cells[i, 27].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[i, 27].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                    namedWorksheet.Cells[i, 28].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[i, 28].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                    namedWorksheet.Cells[i, 29].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[i, 29].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                    //namedWorksheet.Cells[i, 30].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    //namedWorksheet.Cells[i, 30].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                    //namedWorksheet.Cells[i, 31].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    //namedWorksheet.Cells[i, 31].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                    namedWorksheet.Cells[i, 32].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[i, 32].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                    namedWorksheet.Cells[i, 33].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[i, 33].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                    //namedWorksheet.Cells[i, 34].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    //namedWorksheet.Cells[i, 34].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                    namedWorksheet.Cells[i, 35].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[i, 35].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);

                                    namedWorksheet.Cells[i, 62].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[i, 62].Style.Fill.BackgroundColor.SetColor(Color.Blue);
                                    namedWorksheet.Cells[i, 65].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[i, 65].Style.Fill.BackgroundColor.SetColor(Color.Blue);
                                    namedWorksheet.Cells[i, 68].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[i, 68].Style.Fill.BackgroundColor.SetColor(Color.Blue);
                                    namedWorksheet.Cells[i, 71].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[i, 71].Style.Fill.BackgroundColor.SetColor(Color.Blue);
                                    namedWorksheet.Cells[i, 74].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[i, 74].Style.Fill.BackgroundColor.SetColor(Color.Green);
                                    namedWorksheet.Cells[i, 78].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[i, 78].Style.Fill.BackgroundColor.SetColor(Color.Blue);
                                }
                            }
                            //MessageBox.Show(namedWorksheet.Cells[i, 1].Value.ToString());
                            //MessageBox.Show(namedWorksheet.Cells[i + 1, 1].Value.ToString());
                            if(i <dgv3.RowCount)
                            {
                                int j = Convert.ToInt32(namedWorksheet.Cells[i, 1].Value.ToString());
                                int k = Convert.ToInt32(namedWorksheet.Cells[i + 1, 1].Value.ToString());

                                if (j != k)
                                {
                                    //MessageBox.Show(dgv3.ColumnCount.ToString());
                                    for (int l = 1; l <= dgv3.ColumnCount; l++)
                                    {
                                        //MessageBox.Show(namedWorksheet.Cells[i, l].Value.ToString());
                                        //namedWorksheet.Cells[i, l].Style.Border.Top.Style = ExcelBorderStyle.Thick;
                                        //namedWorksheet.Cells[i, l].Style.Border.Left.Style = ExcelBorderStyle.Thick;
                                        //namedWorksheet.Cells[i, l].Style.Border.Right.Style = ExcelBorderStyle.Thick;
                                        //namedWorksheet.InsertRow(i, 1);
                                        //namedWorksheet.Cells[i, l].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
                                    }
                                    //MessageBox.Show(i.ToString());
                                    //namedWorksheet.InsertRow(i, 1);
                                }
                                

                            }


                        }
                        using (ExcelRange Rng = namedWorksheet.Cells[1, 1, 1, dgv3.ColumnCount])
                        {

                            Rng.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            Rng.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            Rng.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            Rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        }
                        int cntRow = dgv3.RowCount;
                        Decimal P = 0;
                        Decimal Q = 0;
                        Decimal R = 0;
                        Decimal S = 0;
                        String t="[0/0]";
                        String v= "[0/0]";
                        for (int s=2; s< cntRow; s++)
                        {
                                String j1 = (String.IsNullOrEmpty(namedWorksheet.Cells[s, 15].Value.ToString())? "dummy" : namedWorksheet.Cells[s, 15].Value.ToString());
                                string k1 = (String.IsNullOrEmpty(namedWorksheet.Cells[s + 1, 15].Value.ToString()) ? "dummy" : namedWorksheet.Cells[s + 1, 15].Value.ToString());
                            //MessageBox.Show(j1.ToString());
                            //MessageBox.Show(k1.ToString());
                            if (j1 == "RED" && k1 == "INC")
                            {
                                    namedWorksheet.InsertRow(s + 1, 1);
                                    namedWorksheet.Cells[s + 1, 15].Value = "ST_RED";
                                namedWorksheet.Cells[s + 1, 30].Value = Convert.ToDecimal(String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[s, 30].Value)) ? "0" : Convert.ToString(namedWorksheet.Cells[s, 30].Value))
                                        + Convert.ToDecimal(String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[s-1, 30].Value)) ? "0" : Convert.ToString(namedWorksheet.Cells[s-1, 30].Value))
                                        + Convert.ToDecimal(String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[s-2, 30].Value)) ? "0" : Convert.ToString(namedWorksheet.Cells[s-2, 30].Value));


                                namedWorksheet.Cells[s + 1, 34].Value = Convert.ToDecimal(String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[s, 34].Value)) ? "0" : Convert.ToString(namedWorksheet.Cells[s, 34].Value))
                                        + Convert.ToDecimal(String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[s - 1, 34].Value)) ? "0" : Convert.ToString(namedWorksheet.Cells[s - 1, 34].Value))
                                        + Convert.ToDecimal(String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[s - 2, 34].Value)) ? "0" : Convert.ToString(namedWorksheet.Cells[s - 2, 34].Value));
                                
                                var t1 = String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[s, 31].Value)) ? "[0/0]" : Convert.ToString(namedWorksheet.Cells[s, 31].Value);
                                var u1 = String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[s - 1, 31].Value)) ? "[0/0]" : Convert.ToString(namedWorksheet.Cells[s - 1, 31].Value);
                                var v1 = String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[s - 2, 31].Value)) ? "[0/0]" : Convert.ToString(namedWorksheet.Cells[s - 2, 31].Value);

                                var Str = "SELECT dbo.fn_CalculateCountSum(" + "'" + t1.ToString().Trim() + "','" + u1.ToString().Trim() + "','" + v1.ToString().Trim() + "'" + ")";

                                using (SqlCommand cmd1 = new SqlCommand(Str, sqlcon))
                                {
                                    var result = cmd1.ExecuteScalar();
                                    namedWorksheet.Cells[s + 1, 31].Value = result;
                                }
                                
                                cntRow++;
                                    //P = Convert.ToDecimal(namedWorksheet.Cells[s + 1, 30].Value);
                            }
                            if (j1 == "INC" && k1 == "RED")
                            {
                                namedWorksheet.InsertRow(s+1, 2);
                                namedWorksheet.Cells[s + 1, 15].Value = "ST_INC";
                                namedWorksheet.Cells[s + 2, 15].Value = "T_RI";
                                namedWorksheet.Cells[s + 1, 30].Value = Convert.ToDecimal(String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[s, 30].Value)) ? "0" : Convert.ToString(namedWorksheet.Cells[s, 30].Value))
                                        + Convert.ToDecimal(String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[s - 1, 30].Value)) ? "0" : Convert.ToString(namedWorksheet.Cells[s - 1, 30].Value))
                                        + Convert.ToDecimal(String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[s - 2, 30].Value)) ? "0" : Convert.ToString(namedWorksheet.Cells[s - 2, 30].Value));

                                namedWorksheet.Cells[s + 1, 34].Value = Convert.ToDecimal(String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[s, 34].Value)) ? "0" : Convert.ToString(namedWorksheet.Cells[s, 34].Value))
                                        + Convert.ToDecimal(String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[s - 1, 34].Value)) ? "0" : Convert.ToString(namedWorksheet.Cells[s - 1, 34].Value))
                                        + Convert.ToDecimal(String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[s - 2, 34].Value)) ? "0" : Convert.ToString(namedWorksheet.Cells[s - 2, 34].Value));

                                var t2 = String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[s, 31].Value)) ? "[0/0]" : Convert.ToString(namedWorksheet.Cells[s, 31].Value);
                                var u2 = String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[s - 1, 31].Value)) ? "[0/0]" : Convert.ToString(namedWorksheet.Cells[s - 1, 31].Value);
                                var v2 = String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[s - 2, 31].Value)) ? "[0/0]" : Convert.ToString(namedWorksheet.Cells[s - 2, 31].Value);

                                var Str1 = "SELECT dbo.fn_CalculateCountSum(" + "'" + t2.ToString().Trim() + "','" + u2.ToString().Trim() + "','" + v2.ToString().Trim() + "'" + ")";

                                using (SqlCommand cmd2 = new SqlCommand(Str1, sqlcon))
                                {
                                    
                                    var result = cmd2.ExecuteScalar();
                                    namedWorksheet.Cells[s + 1, 31].Value = result;
                                    
                                }
                                //cntRow++;
                                cntRow = cntRow + 2;
                                //Q = Convert.ToDecimal(namedWorksheet.Cells[s + 1, 30].Value);
                                //namedWorksheet.Cells[s + 2, 30].Value = (P + Q);
                            }
                            
                        }
                        //MessageBox.Show(cntRow.ToString());
                        namedWorksheet.InsertRow(cntRow + 1, 2);
                        namedWorksheet.Cells[cntRow + 1, 15].Value = "ST_INC";
                        namedWorksheet.Cells[cntRow + 2, 15].Value = "T_RI";
                        namedWorksheet.Cells[cntRow + 1, 30].Value = Convert.ToDecimal(String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[cntRow, 30].Value)) ? "0" : Convert.ToString(namedWorksheet.Cells[cntRow, 30].Value))
                                        + Convert.ToDecimal(String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[cntRow - 1, 30].Value)) ? "0" : Convert.ToString(namedWorksheet.Cells[cntRow - 1, 30].Value))
                                        + Convert.ToDecimal(String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[cntRow - 2, 30].Value)) ? "0" : Convert.ToString(namedWorksheet.Cells[cntRow - 2, 30].Value));

                        namedWorksheet.Cells[cntRow + 1, 34].Value = Convert.ToDecimal(String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[cntRow, 34].Value)) ? "0" : Convert.ToString(namedWorksheet.Cells[cntRow, 34].Value))
                                        + Convert.ToDecimal(String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[cntRow - 1, 34].Value)) ? "0" : Convert.ToString(namedWorksheet.Cells[cntRow - 1, 34].Value))
                                        + Convert.ToDecimal(String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[cntRow - 2, 34].Value)) ? "0" : Convert.ToString(namedWorksheet.Cells[cntRow - 2, 34].Value));


                        var t3 = String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[cntRow, 31].Value)) ? "[0/0]" : Convert.ToString(namedWorksheet.Cells[cntRow, 31].Value);
                        var u3 = String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[cntRow - 1, 31].Value)) ? "[0/0]" : Convert.ToString(namedWorksheet.Cells[cntRow - 1, 31].Value);
                        var v3 = String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[cntRow - 2, 31].Value)) ? "[0/0]" : Convert.ToString(namedWorksheet.Cells[cntRow - 2, 31].Value);

                        var Str3 = "SELECT dbo.fn_CalculateCountSum(" + "'" + t3.ToString().Trim() + "','" + u3.ToString().Trim() + "','" + v3.ToString().Trim() + "'" + ")";

                        using (SqlCommand cmd3 = new SqlCommand(Str3, sqlcon))
                            {
                                var result = cmd3.ExecuteScalar();
                                namedWorksheet.Cells[cntRow + 1, 31].Value = result;
                            }
                            
                        for (int mm=2;mm<= cntRow+2;mm++)
                        {
                            if(Convert.ToString(namedWorksheet.Cells[mm, 15].Value) == "ST_RED")
                            {
                                P = Convert.ToDecimal(namedWorksheet.Cells[mm, 30].Value);
                                R = Convert.ToDecimal(namedWorksheet.Cells[mm, 34].Value);
                                t = namedWorksheet.Cells[mm, 31].Value.ToString();

                                var L_ITEMNO = Convert.ToInt32(namedWorksheet.Cells[mm - 1, 1].Value.ToString());
                                var resultRU_4 ="";
                                var resultRD_4 = "";

                                var StrRU_4 = "SELECT dbo.fnCalculateUPSUM(" + "'" + L_ITEMNO.ToString().Trim() + "U" + "'" + ")";
                                
                                using (SqlCommand cmdRU_4 = new SqlCommand(StrRU_4, sqlcon))
                                {
                                    resultRU_4 = Convert.ToString(cmdRU_4.ExecuteScalar());
                                    namedWorksheet.Cells[mm, 72].Value = resultRU_4;
                                }
                                var StrRD_4 = "SELECT dbo.fnCalculateUPSUM(" + "'" + L_ITEMNO.ToString().Trim() + "D" + "'" + ")";

                                using (SqlCommand cmdRD_4 = new SqlCommand(StrRD_4, sqlcon))
                                {
                                    resultRD_4 = Convert.ToString(cmdRD_4.ExecuteScalar());
                                    namedWorksheet.Cells[mm, 73].Value = resultRD_4;
                                }
                                /*
                                var StrL_5 = "SELECT dbo.fnCalculateUPDWAVG(" + "'" + L_ITEMNO.ToString().Trim() + "U" + "'" + ")";
                                using (SqlCommand cmdL_5 = new SqlCommand(StrL_5, sqlcon))
                                {
                                    var resultL_5 = cmdL_5.ExecuteScalar();
                                    namedWorksheet.Cells[mm, 75].Value = resultL_5;
                                }
                                */
                                String StrL_5 = "SP_CalllingCalculateAVGAMT2NDROW";
                                //MessageBox.Show(var);
                                String ItemNOU = L_ITEMNO.ToString().Trim() + "U";
                                String ItemNOD = L_ITEMNO.ToString().Trim() + "D";
                                SqlCommand cmdL_5 = new SqlCommand();
                                cmdL_5.Connection = sqlcon;
                                cmdL_5.CommandText = StrL_5;
                                cmdL_5.CommandType = CommandType.StoredProcedure;
                                cmdL_5.Parameters.Add("@Var_ITEM_NO", SqlDbType.VarChar).Value = ItemNOU.ToString();
                                //MessageBox.Show("'" + L_ITEMNO.ToString().Trim() + "U" + "'");
                                cmdL_5.CommandTimeout = 3600000;
                                cmdL_5.ExecuteNonQuery();

                                SqlCommand cmll_5 = new SqlCommand("SELECT AVGAMT FROM Result_AVGAMT", sqlcon);
                                string resultL_5 = "";
                                SqlDataReader rdr = cmll_5.ExecuteReader();
                                while (rdr.Read())
                                {
                                    resultL_5 = rdr["AVGAMT"].ToString();
                                }
                                rdr.Close();
                                namedWorksheet.Cells[mm, 75].Value = resultL_5;

                                //next column
                                String StrL_6 = "SP_CalllingCalculateAVGAMT2NDROW";
                                SqlCommand cmdL_6 = new SqlCommand();
                                cmdL_6.Connection = sqlcon;
                                cmdL_6.CommandText = StrL_6;
                                cmdL_6.CommandType = CommandType.StoredProcedure;
                                cmdL_6.Parameters.Add("@Var_ITEM_NO", SqlDbType.VarChar).Value = ItemNOD.ToString();
                                //MessageBox.Show("'" + L_ITEMNO.ToString().Trim() + "U" + "'");
                                cmdL_6.CommandTimeout = 3600000;
                                cmdL_6.ExecuteNonQuery();

                                SqlCommand cmll_6 = new SqlCommand("SELECT AVGAMT FROM Result_AVGAMT", sqlcon);
                                string resultL_6 = "";
                                SqlDataReader rdr6 = cmll_6.ExecuteReader();
                                while (rdr6.Read())
                                {
                                    resultL_6 = rdr6["AVGAMT"].ToString();
                                }
                                rdr6.Close();
                                namedWorksheet.Cells[mm, 76].Value = resultL_6;

                                
                                var StrUDIRES = Convert.ToDecimal(namedWorksheet.Cells[mm - 1, 37].Value);
                                //MessageBox.Show(StrUDIRES.ToString());
                                //MessageBox.Show(resultRU_4.ToString());
                                //MessageBox.Show(Convert.ToInt32((resultRU_4.Replace("UP", "")).Replace("DW", "")).ToString());

                                if((( Convert.ToInt32((resultRU_4.Replace(@"UP/","")).Replace(@"DW/",""))> Convert.ToInt32((resultRD_4.Replace(@"UP/", "")).Replace(@"DW/", "")) ) && StrUDIRES>0) )
                                {
                                    namedWorksheet.Cells[mm, 77].Value = "C";
                                }
                                else if (((Convert.ToInt32((resultRU_4.Replace(@"UP/", "")).Replace(@"DW/", "")) < Convert.ToInt32((resultRD_4.Replace(@"UP/", "")).Replace(@"DW/", ""))) && StrUDIRES < 0))
                                {
                                    namedWorksheet.Cells[mm, 77].Value = "C";
                                }
                                else
                                {
                                    namedWorksheet.Cells[mm, 77].Value = "W";
                                }
                                
                                namedWorksheet.Cells[mm, 62].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[mm, 62].Style.Fill.BackgroundColor.SetColor(Color.Blue);
                                namedWorksheet.Cells[mm, 65].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[mm, 65].Style.Fill.BackgroundColor.SetColor(Color.Blue);
                                namedWorksheet.Cells[mm, 68].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[mm, 68].Style.Fill.BackgroundColor.SetColor(Color.Blue);
                                namedWorksheet.Cells[mm, 71].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[mm, 71].Style.Fill.BackgroundColor.SetColor(Color.Blue);
                                namedWorksheet.Cells[mm, 74].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[mm, 74].Style.Fill.BackgroundColor.SetColor(Color.Green);
                                namedWorksheet.Cells[mm, 78].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[mm, 78].Style.Fill.BackgroundColor.SetColor(Color.Blue);
                            }
                            if (Convert.ToString(namedWorksheet.Cells[mm, 15].Value) == "ST_INC")
                            {
                                Q = Convert.ToDecimal(namedWorksheet.Cells[mm, 30].Value);
                                S = Convert.ToDecimal(namedWorksheet.Cells[mm, 34].Value);
                                v = namedWorksheet.Cells[mm, 31].Value.ToString();

                                namedWorksheet.Cells[mm, 62].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[mm, 62].Style.Fill.BackgroundColor.SetColor(Color.Blue);
                                namedWorksheet.Cells[mm, 65].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[mm, 65].Style.Fill.BackgroundColor.SetColor(Color.Blue);
                                namedWorksheet.Cells[mm, 68].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[mm, 68].Style.Fill.BackgroundColor.SetColor(Color.Blue);
                                namedWorksheet.Cells[mm, 71].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[mm, 71].Style.Fill.BackgroundColor.SetColor(Color.Blue);
                                namedWorksheet.Cells[mm, 74].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[mm, 74].Style.Fill.BackgroundColor.SetColor(Color.Green);
                                namedWorksheet.Cells[mm, 78].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[mm, 78].Style.Fill.BackgroundColor.SetColor(Color.Blue);
                            }
                            if (Convert.ToString(namedWorksheet.Cells[mm, 15].Value) == "T_RI")
                            {
                                namedWorksheet.Cells[mm, 62].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[mm, 62].Style.Fill.BackgroundColor.SetColor(Color.Blue);
                                namedWorksheet.Cells[mm, 65].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[mm, 65].Style.Fill.BackgroundColor.SetColor(Color.Blue);
                                namedWorksheet.Cells[mm, 68].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[mm, 68].Style.Fill.BackgroundColor.SetColor(Color.Blue);
                                namedWorksheet.Cells[mm, 71].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[mm, 71].Style.Fill.BackgroundColor.SetColor(Color.Blue);
                                namedWorksheet.Cells[mm, 74].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[mm, 74].Style.Fill.BackgroundColor.SetColor(Color.Green);
                                namedWorksheet.Cells[mm, 78].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[mm, 78].Style.Fill.BackgroundColor.SetColor(Color.Blue);

                                for (int l = 1; l <= dgv3.ColumnCount; l++)
                                {
                                    //namedWorksheet.Cells[mm, l].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    //namedWorksheet.Cells[mm, l].Style.Fill.BackgroundColor.SetColor(Color.Red);

                                    namedWorksheet.Cells[mm, l].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
                                    namedWorksheet.Cells[mm, l].Style.Border.Bottom.Color.SetColor(Color.Red);
                                }
                                    namedWorksheet.Cells[mm, 30].Value = (P + Q);
                                    namedWorksheet.Cells[mm, 34].Value = (R + S);
                                
                                var Str4 = "SELECT dbo.fn_CalculateCountSum(" + "'" + t.ToString().Trim() + "','" + v.ToString().Trim() + "','" + ("[0/0]").ToString().Trim() + "'" + ")";

                                using (SqlCommand cmd4 = new SqlCommand(Str4, sqlcon))
                                {
                                    var result = cmd4.ExecuteScalar();
                                    namedWorksheet.Cells[mm, 31].Value = result;
                                }
                            }

                        }
                        for (int mn = 2; mn <= cntRow + 2; mn++)
                        {
                            if (Convert.ToString(namedWorksheet.Cells[mn, 15].Value) == "T_RI")
                            {
                                //MessageBox.Show(namedWorksheet.Cells[mn-2, 1].Value.ToString());
                                var L_ITEMNO = Convert.ToInt32(namedWorksheet.Cells[mn - 2, 1].Value.ToString());
                                
                                var StrL5 = "SELECT dbo.fn_Calculateupsum(" + "'" + L_ITEMNO.ToString().Trim() +"U" + "'" + ")";

                                using (SqlCommand cmdL5 = new SqlCommand(StrL5, sqlcon))
                                {
                                    var resultL5 = cmdL5.ExecuteScalar();
                                    namedWorksheet.Cells[mn, 66].Value = resultL5;
                                }

                                var StrL6 = "SELECT dbo.fn_Calculateupsum(" + "'" + L_ITEMNO.ToString().Trim() + "D" + "'" + ")";
                                using (SqlCommand cmdL6 = new SqlCommand(StrL6, sqlcon))
                                {
                                    var resultL6 = cmdL6.ExecuteScalar();
                                    namedWorksheet.Cells[mn, 67].Value = resultL6;
                                }

                                var StrL7 = "SELECT dbo.fn_CalculateupCount(" + "'" + L_ITEMNO.ToString().Trim() + "U" + "'" + ")";
                                using (SqlCommand cmdL7 = new SqlCommand(StrL7, sqlcon))
                                {
                                    var resultL7 = cmdL7.ExecuteScalar();
                                    namedWorksheet.Cells[mn, 69].Value = resultL7;
                                }

                                var StrL8 = "SELECT dbo.fn_CalculateupCount(" + "'" + L_ITEMNO.ToString().Trim() + "D" + "'" + ")";
                                using (SqlCommand cmdL8 = new SqlCommand(StrL8, sqlcon))
                                {
                                    var resultL8 = cmdL8.ExecuteScalar();
                                    namedWorksheet.Cells[mn, 70].Value = resultL8;
                                }

                                var T_Value = Convert.ToDecimal(namedWorksheet.Cells[mn, 30].Value.ToString());
                                var T_3COLValue = Convert.ToDecimal(namedWorksheet.Cells[mn, 34].Value.ToString());
                                var T_CNTPN = namedWorksheet.Cells[mn, 31].Value.ToString();
                                //MessageBox.Show(namedWorksheet.Cells[mn, 31].Value.ToString());
                                //MessageBox.Show(namedWorksheet.Cells[mn - 1, 37].Value.ToString());
                                namedWorksheet.Cells[mn, 37].Value = Convert.ToDecimal(String.IsNullOrEmpty(namedWorksheet.Cells[mn - 2, 37].Value.ToString()) ? "0" : namedWorksheet.Cells[mn - 2, 37].Value.ToString());
                                namedWorksheet.Cells[mn, 38].Value = Convert.ToDecimal(String.IsNullOrEmpty(namedWorksheet.Cells[mn - 2, 38].Value.ToString()) ? "0" : namedWorksheet.Cells[mn - 2, 38].Value.ToString());

                                var T_IRES = Convert.ToDecimal(String.IsNullOrEmpty(namedWorksheet.Cells[mn - 2, 37].Value.ToString()) ? "0" : namedWorksheet.Cells[mn - 2, 37].Value.ToString());
                                
                                var Str5 = "SELECT dbo.fn_FinalCountPN(" + "'" + T_CNTPN.ToString().Trim() + "'"+ ")";

                                using (SqlCommand cmd5 = new SqlCommand(Str5, sqlcon))
                                {
                                    var result5 = cmd5.ExecuteScalar();
                                    //namedWorksheet.Cells[cntRow + 1, 31].Value = result;

                                    if (T_Value > 0)
                                    {
                                        if (Convert.ToInt32(result5.ToString()) == 1)
                                        {
                                            if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'C';
                                            }
                                            else if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'W';
                                            }

                                        }
                                        if (Convert.ToInt32(result5.ToString()) == 0)
                                        {
                                            if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'W';
                                            }
                                            else if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'C';
                                            }

                                        }
                                        if (Convert.ToInt32(result5.ToString()) == 2)
                                        {
                                            if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'C';
                                            }
                                            else if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'W';
                                            }

                                        }
                                    }
                                    else if (T_Value < 0)
                                    {
                                        if (Convert.ToInt32(result5.ToString()) == 0)
                                        {
                                            if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'C';
                                            }
                                            else if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'W';
                                            }

                                        }
                                        if (Convert.ToInt32(result5.ToString()) == 1)
                                        {
                                            if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'W';
                                            }
                                            else if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'C';
                                            }

                                        }
                                        if (Convert.ToInt32(result5.ToString()) == 2)
                                        {
                                            if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'C';
                                            }
                                            else if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'W';
                                            }

                                        }
                                    }


                                    //3col

                                    if (T_3COLValue > 0)
                                    {
                                        if (Convert.ToInt32(result5.ToString()) == 1)
                                        {
                                            if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'C';
                                            }
                                            else if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'W';
                                            }

                                        }
                                        if (Convert.ToInt32(result5.ToString()) == 0)
                                        {
                                            if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'W';
                                            }
                                            else if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'C';
                                            }

                                        }
                                        if (Convert.ToInt32(result5.ToString()) == 2)
                                        {
                                            if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'C';
                                            }
                                            else if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'W';
                                            }

                                        }
                                    }
                                    else if (T_3COLValue < 0)
                                    {
                                        if (Convert.ToInt32(result5.ToString()) == 0)
                                        {
                                            if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'C';
                                            }
                                            else if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'W';
                                            }

                                        }
                                        if (Convert.ToInt32(result5.ToString()) == 1)
                                        {
                                            if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'W';
                                            }
                                            else if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'C';
                                            }

                                        }
                                        if (Convert.ToInt32(result5.ToString()) == 2)
                                        {
                                            if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'C';
                                            }
                                            else if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'W';
                                            }

                                        }
                                    }
                                   
                                }
                                
                            }

                            if(Convert.ToString(namedWorksheet.Cells[mn, 15].Value) == "ST_RED")
                            {
                                var T_Value = Convert.ToDecimal(namedWorksheet.Cells[mn, 30].Value.ToString());
                                var T_3COLValue = Convert.ToDecimal(namedWorksheet.Cells[mn, 34].Value.ToString());
                                var T_CNTPN = namedWorksheet.Cells[mn, 31].Value.ToString();
                                //MessageBox.Show(namedWorksheet.Cells[mn, 31].Value.ToString());
                                //MessageBox.Show(namedWorksheet.Cells[mn - 1, 37].Value.ToString());
                                namedWorksheet.Cells[mn, 37].Value = Convert.ToDecimal(String.IsNullOrEmpty(namedWorksheet.Cells[mn - 1, 37].Value.ToString()) ? "0" : namedWorksheet.Cells[mn - 1, 37].Value.ToString());
                                namedWorksheet.Cells[mn, 38].Value = Convert.ToDecimal(String.IsNullOrEmpty(namedWorksheet.Cells[mn - 1, 38].Value.ToString()) ? "0" : namedWorksheet.Cells[mn - 1, 38].Value.ToString());
                                var T_IRES = Convert.ToDecimal(String.IsNullOrEmpty(namedWorksheet.Cells[mn - 1, 37].Value.ToString()) ? "0" : namedWorksheet.Cells[mn - 1, 37].Value.ToString());

                                var Str6 = "SELECT dbo.fn_FinalCountPN(" + "'" + T_CNTPN.ToString().Trim() + "'" + ")";

                                using (SqlCommand cmd6 = new SqlCommand(Str6, sqlcon))
                                {
                                    var result6 = cmd6.ExecuteScalar();
                                    //namedWorksheet.Cells[cntRow + 1, 31].Value = result;
                                    if (T_Value > 0)
                                    {
                                        if (Convert.ToInt32(result6.ToString()) == 1)
                                        {
                                            if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'C';
                                            }
                                            else if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'W';
                                            }

                                        }
                                        if (Convert.ToInt32(result6.ToString()) == 0)
                                        {
                                            if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'W';
                                            }
                                            else if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'C';
                                            }

                                        }
                                        if (Convert.ToInt32(result6.ToString()) == 2)
                                        {
                                            if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'C';
                                            }
                                            else if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'W';
                                            }

                                        }
                                    }
                                    else if (T_Value < 0)
                                    {
                                        if (Convert.ToInt32(result6.ToString()) == 0)
                                        {
                                            if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'C';
                                            }
                                            else if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'W';
                                            }

                                        }
                                        if (Convert.ToInt32(result6.ToString()) == 1)
                                        {
                                            if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'W';
                                            }
                                            else if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'C';
                                            }

                                        }
                                        if (Convert.ToInt32(result6.ToString()) == 2)
                                        {
                                            if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'C';
                                            }
                                            else if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'W';
                                            }

                                        }
                                    }

                                
                                    //3col

                                    if (T_3COLValue > 0)
                                    {
                                        if (Convert.ToInt32(result6.ToString()) == 1)
                                        {
                                            if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'C';
                                            }
                                            else if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'W';
                                            }

                                        }
                                        if (Convert.ToInt32(result6.ToString()) == 0)
                                        {
                                            if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'W';
                                            }
                                            else if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'C';
                                            }

                                        }
                                        if (Convert.ToInt32(result6.ToString()) == 2)
                                        {
                                            if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'C';
                                            }
                                            else if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'W';
                                            }

                                        }
                                    }
                                    else if (T_3COLValue < 0)
                                    {
                                        if (Convert.ToInt32(result6.ToString()) == 0)
                                        {
                                            if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'C';
                                            }
                                            else if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'W';
                                            }

                                        }
                                        if (Convert.ToInt32(result6.ToString()) == 1)
                                        {
                                            if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'W';
                                            }
                                            else if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'C';
                                            }

                                        }
                                        if (Convert.ToInt32(result6.ToString()) == 2)
                                        {
                                            if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'C';
                                            }
                                            else if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'W';
                                            }

                                        }
                                    }


                                }

                                //NEWLY ADDED COLUMN
                                namedWorksheet.Cells[mn, 52].Value = namedWorksheet.Cells[mn - 1, 52].Value;
                                //MessageBox.Show((Convert.ToDecimal(String.IsNullOrEmpty(namedWorksheet.Cells[mn, 34].Value.ToString()) ? "0" : namedWorksheet.Cells[mn, 34].Value.ToString())).ToString());
                                //MessageBox.Show((Convert.ToDecimal(String.IsNullOrEmpty(namedWorksheet.Cells[mn + 4, 34].Value.ToString()) ? "0" : namedWorksheet.Cells[mn + 4, 34].Value.ToString())).ToString());
                                namedWorksheet.Cells[mn, 53].Value = Math.Abs((Convert.ToDecimal(String.IsNullOrEmpty(namedWorksheet.Cells[mn, 34].Value.ToString()) ? "0" : namedWorksheet.Cells[mn, 34].Value.ToString())) - (Convert.ToDecimal(String.IsNullOrEmpty(namedWorksheet.Cells[mn + 4, 34].Value.ToString()) ? "0" : namedWorksheet.Cells[mn + 4, 34].Value.ToString())));
                                var Str1_6 = "SELECT dbo.fn_Calculatecntersum(" + "'" + namedWorksheet.Cells[mn - 1, 7].Value.ToString().Trim() + "'" + ",'" + namedWorksheet.Cells[mn - 1, 9].Value.ToString().Trim() + "'," + "'" + namedWorksheet.Cells[mn - 1, 11].Value.ToString().Trim() + "'" + ")";

                                using (SqlCommand cmd1_6 = new SqlCommand(Str1_6, sqlcon))
                                {
                                    var result1_6 = cmd1_6.ExecuteScalar();
                                    //MessageBox.Show(result6.ToString());
                                    if (Convert.ToInt32(result1_6.ToString()) < 0)
                                    {
                                        namedWorksheet.Cells[mn, 54].Value = Convert.ToInt32(result1_6.ToString());
                                    }
                                    else
                                    {
                                        namedWorksheet.Cells[mn, 54].Value = 0;
                                    }
                                }
                                var Str1_7 = "SELECT dbo.fn_Calculatecntersum(" + "'" + namedWorksheet.Cells[mn - 1, 12].Value.ToString().Trim() + "'" + ",'" + namedWorksheet.Cells[mn - 1, 13].Value.ToString().Trim() + "'," + "'" + namedWorksheet.Cells[mn - 1, 14].Value.ToString().Trim() + "'" + ")";

                                using (SqlCommand cmd1_7 = new SqlCommand(Str1_7, sqlcon))
                                {
                                    var result1_7 = cmd1_7.ExecuteScalar();
                                    //MessageBox.Show(result6.ToString());
                                    if (Convert.ToInt32(result1_7.ToString()) < 0)
                                    {
                                        namedWorksheet.Cells[mn, 55].Value = Convert.ToInt32(result1_7.ToString());
                                    }
                                    else
                                    {
                                        namedWorksheet.Cells[mn, 55].Value = 0;
                                    }
                                }

                                var Str1_8 = "SELECT dbo.fn_Calculatecntersum(" + "'" + namedWorksheet.Cells[mn - 1, 40].Value.ToString().Trim() + "'" + ",'" + namedWorksheet.Cells[mn - 1, 42].Value.ToString().Trim() + "'," + "'" + namedWorksheet.Cells[mn - 1, 44].Value.ToString().Trim() + "'" + ")";

                                using (SqlCommand cmd1_8 = new SqlCommand(Str1_8, sqlcon))
                                {
                                    var result1_8 = cmd1_8.ExecuteScalar();
                                    //MessageBox.Show(result6.ToString());
                                    if (Convert.ToInt32(result1_8.ToString()) < 0)
                                    {
                                        namedWorksheet.Cells[mn, 56].Value = Convert.ToInt32(result1_8.ToString());
                                    }
                                    else
                                    {
                                        namedWorksheet.Cells[mn, 56].Value = 0;
                                    }
                                }

                                namedWorksheet.Cells[mn, 57].Value = Convert.ToInt32(namedWorksheet.Cells[mn, 54].Value.ToString()) + Convert.ToInt32(namedWorksheet.Cells[mn, 55].Value.ToString()) + Convert.ToInt32(namedWorksheet.Cells[mn, 56].Value.ToString());

                                        //Calculation for column [NO-OF-Ls/NO-OF-Hs]
                                        var result1_9 = "";
                                        var result1_10 = "";
                                        var result1_11 = "";
                                        var result1_12 = "";

                                        var Str1_9 = "SELECT dbo.fn_CalculateCountSum(" + "'" + namedWorksheet.Cells[mn - 1, 7].Value.ToString().Trim() + "'" + ",'" + namedWorksheet.Cells[mn - 1, 9].Value.ToString().Trim() + "'," + "'" + namedWorksheet.Cells[mn - 1, 11].Value.ToString().Trim() + "'" + ")";

                                        using (SqlCommand cmd1_9 = new SqlCommand(Str1_9, sqlcon))
                                        {
                                            result1_9 = Convert.ToString(cmd1_9.ExecuteScalar());
                                        }
                                        var Str1_10 = "SELECT dbo.fn_CalculateCountSum(" + "'" + namedWorksheet.Cells[mn - 1, 12].Value.ToString().Trim() + "'" + ",'" + namedWorksheet.Cells[mn - 1, 13].Value.ToString().Trim() + "'," + "'" + namedWorksheet.Cells[mn - 1, 14].Value.ToString().Trim() + "'" + ")";

                                        using (SqlCommand cmd1_10 = new SqlCommand(Str1_10, sqlcon))
                                        {
                                            result1_10 = Convert.ToString(cmd1_10.ExecuteScalar());
                                        }
                                        var Str1_11 = "SELECT dbo.fn_CalculateCountSum(" + "'" + namedWorksheet.Cells[mn - 1, 40].Value.ToString().Trim() + "'" + ",'" + namedWorksheet.Cells[mn - 1, 42].Value.ToString().Trim() + "'," + "'" + namedWorksheet.Cells[mn - 1, 44].Value.ToString().Trim() + "'" + ")";

                                        using (SqlCommand cmd1_11 = new SqlCommand(Str1_11, sqlcon))
                                        {
                                            result1_11 = Convert.ToString(cmd1_11.ExecuteScalar());
                                        }
                                        var Str1_12 = "SELECT dbo.fn_CalculateCountSum(" + "'" + result1_9.ToString().Trim() + "'" + ",'" + result1_10.ToString().Trim() + "'," + "'" + result1_11.ToString().Trim() + "'" + ")";

                                        using (SqlCommand cmd1_12 = new SqlCommand(Str1_12, sqlcon))
                                        {
                                            result1_12 = Convert.ToString(cmd1_12.ExecuteScalar());
                                        }

                                        namedWorksheet.Cells[mn, 58].Value = result1_12.ToString();
                                        //End of Calculation for column [NO-OF-Ls/NO-OF-Hs]
                            }
                            if (Convert.ToString(namedWorksheet.Cells[mn, 15].Value) == "ST_INC")
                            {
                                var T_Value = Convert.ToDecimal(namedWorksheet.Cells[mn, 30].Value.ToString());
                                var T_3COLValue = Convert.ToDecimal(namedWorksheet.Cells[mn, 34].Value.ToString());
                                var T_CNTPN = namedWorksheet.Cells[mn, 31].Value.ToString();
                                //MessageBox.Show(namedWorksheet.Cells[mn, 31].Value.ToString());
                                //MessageBox.Show(namedWorksheet.Cells[mn - 1, 37].Value.ToString());
                                namedWorksheet.Cells[mn, 37].Value = Convert.ToDecimal(String.IsNullOrEmpty(namedWorksheet.Cells[mn - 1, 37].Value.ToString()) ? "0" : namedWorksheet.Cells[mn - 1, 37].Value.ToString());
                                namedWorksheet.Cells[mn, 38].Value = Convert.ToDecimal(String.IsNullOrEmpty(namedWorksheet.Cells[mn - 1, 38].Value.ToString()) ? "0" : namedWorksheet.Cells[mn - 1, 38].Value.ToString());

                                var T_IRES = Convert.ToDecimal(String.IsNullOrEmpty(namedWorksheet.Cells[mn - 1, 37].Value.ToString()) ? "0" : namedWorksheet.Cells[mn - 1, 37].Value.ToString());

                                var Str6 = "SELECT dbo.fn_FinalCountPN(" + "'" + T_CNTPN.ToString().Trim() + "'" + ")";

                                using (SqlCommand cmd6 = new SqlCommand(Str6, sqlcon))
                                {
                                    var result6 = cmd6.ExecuteScalar();
                                    //namedWorksheet.Cells[cntRow + 1, 31].Value = result;

                                    if (T_Value > 0)
                                    {
                                        if (Convert.ToInt32(result6.ToString()) == 1)
                                        {
                                            if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'C';
                                            }
                                            else if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'W';
                                            }

                                        }
                                        if (Convert.ToInt32(result6.ToString()) == 0)
                                        {
                                            if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'W';
                                            }
                                            else if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'C';
                                            }

                                        }
                                        if (Convert.ToInt32(result6.ToString()) == 2)
                                        {
                                            if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'C';
                                            }
                                            else if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'W';
                                            }

                                        }

                                    }
                                    else if (T_Value < 0)
                                    {
                                        if (Convert.ToInt32(result6.ToString()) == 0)
                                        {
                                            if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'C';
                                            }
                                            else if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'W';
                                            }

                                        }
                                        if (Convert.ToInt32(result6.ToString()) == 1)
                                        {
                                            if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'W';
                                            }
                                            else if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'C';
                                            }

                                        }
                                        if (Convert.ToInt32(result6.ToString()) == 2)
                                        {
                                            if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'C';
                                            }
                                            else if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'W';
                                            }

                                        }
                                    }


                                    //3col

                                    if (T_3COLValue > 0)
                                    {
                                        if (Convert.ToInt32(result6.ToString()) == 1)
                                        {
                                            if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'C';
                                            }
                                            else if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'W';
                                            }

                                        }
                                        if (Convert.ToInt32(result6.ToString()) == 0)
                                        {
                                            if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'W';
                                            }
                                            else if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'C';
                                            }

                                        }
                                        if (Convert.ToInt32(result6.ToString()) == 2)
                                        {
                                            if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'C';
                                            }
                                            else if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'W';
                                            }

                                        }
                                    }
                                    else if (T_3COLValue < 0)
                                    {
                                        if (Convert.ToInt32(result6.ToString()) == 0)
                                        {
                                            if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'C';
                                            }
                                            else if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'W';
                                            }

                                        }
                                        if (Convert.ToInt32(result6.ToString()) == 1)
                                        {
                                            if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'W';
                                            }
                                            else if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'C';
                                            }

                                        }
                                        if (Convert.ToInt32(result6.ToString()) == 2)
                                        {
                                            if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'C';
                                            }
                                            else if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'W';
                                            }

                                        }
                                    }

                                }
                                //newly added 
                                    namedWorksheet.Cells[mn, 52].Value = namedWorksheet.Cells[mn - 1, 52].Value;
                                    namedWorksheet.Cells[mn, 53].Value = Math.Abs((Convert.ToDecimal(String.IsNullOrEmpty(namedWorksheet.Cells[mn, 34].Value.ToString()) ? "0" : namedWorksheet.Cells[mn, 34].Value.ToString())) - (Convert.ToDecimal(String.IsNullOrEmpty(namedWorksheet.Cells[mn - 4, 34].Value.ToString()) ? "0" : namedWorksheet.Cells[mn - 4, 34].Value.ToString())));
                                    //var s1diff = namedWorksheet.Cells[mn - 1, 7].Value.ToString()
                                    var Str_6 = "SELECT dbo.fn_Calculatecntersum(" + "'" + namedWorksheet.Cells[mn - 1, 7].Value.ToString().Trim() + "'" + ",'" + namedWorksheet.Cells[mn - 1, 9].Value.ToString().Trim() + "',"+ "'" + namedWorksheet.Cells[mn - 1, 11].Value.ToString().Trim() + "'" + ")";

                                    using (SqlCommand cmd_6 = new SqlCommand(Str_6, sqlcon))
                                    {
                                        var result_6 = cmd_6.ExecuteScalar();
                                        //MessageBox.Show(result6.ToString());
                                        if(Convert.ToInt32(result_6.ToString())<0)
                                        {
                                            namedWorksheet.Cells[mn, 54].Value = Convert.ToInt32(result_6.ToString());
                                        }
                                        else
                                        {
                                            namedWorksheet.Cells[mn, 54].Value = 0;
                                        }
                                    }
                                var Str_7 = "SELECT dbo.fn_Calculatecntersum(" + "'" + namedWorksheet.Cells[mn - 1, 12].Value.ToString().Trim() + "'" + ",'" + namedWorksheet.Cells[mn - 1, 13].Value.ToString().Trim() + "'," + "'" + namedWorksheet.Cells[mn - 1, 14].Value.ToString().Trim() + "'" + ")";

                                using (SqlCommand cmd_7 = new SqlCommand(Str_7, sqlcon))
                                {
                                    var result_7 = cmd_7.ExecuteScalar();
                                    //MessageBox.Show(result6.ToString());
                                    if (Convert.ToInt32(result_7.ToString()) < 0)
                                    {
                                        namedWorksheet.Cells[mn, 55].Value = Convert.ToInt32(result_7.ToString());
                                    }
                                    else
                                    {
                                        namedWorksheet.Cells[mn, 55].Value = 0;
                                    }
                                }
                                
                                var Str_8 = "SELECT dbo.fn_Calculatecntersum(" + "'" + namedWorksheet.Cells[mn - 1, 40].Value.ToString().Trim() + "'" + ",'" + namedWorksheet.Cells[mn - 1, 42].Value.ToString().Trim() + "'," + "'" + namedWorksheet.Cells[mn - 1, 44].Value.ToString().Trim() + "'" + ")";

                                using (SqlCommand cmd_8 = new SqlCommand(Str_8, sqlcon))
                                {
                                    var result_8 = cmd_8.ExecuteScalar();
                                    //MessageBox.Show(result6.ToString());
                                    if (Convert.ToInt32(result_8.ToString()) < 0)
                                    {
                                        namedWorksheet.Cells[mn, 56].Value = Convert.ToInt32(result_8.ToString());
                                    }
                                    else
                                    {
                                        namedWorksheet.Cells[mn, 56].Value = 0;
                                    }
                                }

                                namedWorksheet.Cells[mn, 57].Value = Convert.ToInt32(namedWorksheet.Cells[mn, 54].Value.ToString()) + Convert.ToInt32(namedWorksheet.Cells[mn, 55].Value.ToString()) + Convert.ToInt32(namedWorksheet.Cells[mn, 56].Value.ToString());

                                //END 
                                //Calculation for column [NO-OF-Ls/NO-OF-Hs]
                                var result_9 = "";
                                var result_10 = "";
                                var result_11 = "";
                                var result_12 = "";

                                var Str_9 = "SELECT dbo.fn_CalculateCountSum(" + "'" + namedWorksheet.Cells[mn - 1, 7].Value.ToString().Trim() + "'" + ",'" + namedWorksheet.Cells[mn - 1, 9].Value.ToString().Trim() + "'," + "'" + namedWorksheet.Cells[mn - 1, 11].Value.ToString().Trim() + "'" + ")";

                                using (SqlCommand cmd_9 = new SqlCommand(Str_9, sqlcon))
                                {
                                    result_9 = Convert.ToString(cmd_9.ExecuteScalar());
                                }
                                var Str_10 = "SELECT dbo.fn_CalculateCountSum(" + "'" + namedWorksheet.Cells[mn - 1, 12].Value.ToString().Trim() + "'" + ",'" + namedWorksheet.Cells[mn - 1, 13].Value.ToString().Trim() + "'," + "'" + namedWorksheet.Cells[mn - 1, 14].Value.ToString().Trim() + "'" + ")";

                                using (SqlCommand cmd_10 = new SqlCommand(Str_10, sqlcon))
                                {
                                     result_10 = Convert.ToString(cmd_10.ExecuteScalar());
                                }
                                var Str_11 = "SELECT dbo.fn_CalculateCountSum(" + "'" + namedWorksheet.Cells[mn - 1, 40].Value.ToString().Trim() + "'" + ",'" + namedWorksheet.Cells[mn - 1, 42].Value.ToString().Trim() + "'," + "'" + namedWorksheet.Cells[mn - 1, 44].Value.ToString().Trim() + "'" + ")";

                                using (SqlCommand cmd_11 = new SqlCommand(Str_11, sqlcon))
                                {
                                     result_11 = Convert.ToString(cmd_11.ExecuteScalar());
                                }
                                var Str_12 = "SELECT dbo.fn_CalculateCountSum(" + "'" + result_9.ToString().Trim() + "'" + ",'" + result_10.ToString().Trim() + "'," + "'" + result_11.ToString().Trim() + "'" + ")";

                                using (SqlCommand cmd_12 = new SqlCommand(Str_12, sqlcon))
                                {
                                     result_12 = Convert.ToString(cmd_12.ExecuteScalar());
                                }

                                namedWorksheet.Cells[mn, 58].Value = result_12.ToString();

                            }

                        }
                            //Save your file
                            excelPackage.Save();
                    }
                
                }
                catch(Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                    //wb.Save();
                    //wb.Close(true);
                    //excelApp.Quit();
                    //System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                }
                finally
                {
                    //wb.Save();
                    //wb.Close(true);
                    //excelApp.Quit();
                    //System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                    //sqlcon.Close();
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
        private void Main_Load(object sender, EventArgs e)
        {
            this.Refresh();
            this.Show();
            this.Location = new Point(0, 0);
            this.Size = Screen.PrimaryScreen.WorkingArea.Size;

            try
            {
                sqlcon.Open();
                SqlDataAdapter sda = new SqlDataAdapter("SELECT sum([C_Count]) as C_Total,sum([W_Count]) as W_Total FROM [dbo].[GraphPercentageCalc] WHERE [RED/INC] = 'INC'", sqlcon);
                DataTable dt = new DataTable();
                sda.Fill(dt);
                dgGraphPercntg.DataSource = dt;

                dgv3.DataSource = null;
                /*
                ===============================
                This are is later decided to be replaced with Grid
                ================================
                chart1.Titles.Add("Increasing Accuracy Percentage");
                for (int index = 0; index < dt.Rows.Count; index++)
                {
                    //String Plabel = dt.Rows[0]["C_Total"].ToString();
                    int CorrectCount = Convert.ToInt32(dt.Rows[0]["C_Total"].ToString());
                    int WrongCount = Convert.ToInt32(dt.Rows[0]["W_Total"].ToString());
                    chart1.Series["S1"].Points.AddXY("INC_C", CorrectCount);
                    chart1.Series["S1"].Points.AddXY("INC_W", WrongCount);
                }
                */
                SqlDataAdapter sda1 = new SqlDataAdapter("SELECT sum([C_Count]) as C_Total,sum([W_Count]) as W_Total FROM [dbo].[GraphPercentageCalc] WHERE [RED/INC] = 'RED'", sqlcon);
                DataTable dt1 = new DataTable();
                sda1.Fill(dt1);
                chart2.Titles.Add("Reducing Accuracy Percentage");
                for (int index1 = 0; index1 < dt.Rows.Count; index1++)
                {
                    //String Plabel = dt.Rows[0]["C_Total"].ToString();
                    int CorrectCount1 = Convert.ToInt32(dt1.Rows[0]["C_Total"].ToString());
                    int WrongCount1 = Convert.ToInt32(dt1.Rows[0]["W_Total"].ToString());
                    chart2.Series["S2"].Points.AddXY("RED_C", CorrectCount1);
                    chart2.Series["S2"].Points.AddXY("RED_W", WrongCount1);
                }
                




            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
            finally
            {
                sqlcon.Close();
            }
            
        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void btnEditInput_Click(object sender, EventArgs e)
        {
            this.Close();
            EditInputData ed = new EditInputData();
            ed.Show();
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void lblUD3_Click(object sender, EventArgs e)
        {

        }

        private void btnBrwsFl_Click(object sender, EventArgs e)
        {
            //To where your opendialog box get starting location. My initial directory location is desktop.
            openFileDialog1.InitialDirectory = "C://Desktop";
            //Your opendialog box title name.
            openFileDialog1.Title = "Select file to be upload.";
            //which type file format you want to upload in database. just add them.
            openFileDialog1.Filter = "Select Valid Document(*.pdf; *.doc; *.xlsx; *.html)|*.pdf; *.docx; *.xlsx; *.html";
            //FilterIndex property represents the index of the filter currently selected in the file dialog box.
            openFileDialog1.FilterIndex = 1;
            try
            {
                if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    if (openFileDialog1.CheckFileExists)
                    {
                        string path = System.IO.Path.GetFullPath(openFileDialog1.FileName);
                        //label1.Text = path;
                    }
                }
                else
                {
                    MessageBox.Show("Please Upload document.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btlUpldDoc_Click(object sender, EventArgs e)
        {
            string conString = string.Empty;
            string extension = Path.GetExtension(openFileDialog1.FileName);
            string excelPath = System.IO.Path.GetFullPath(openFileDialog1.FileName);
            if (excelPath.Contains("PP-FILE-WITH DATE - V2.xlsx"))
            {
                try
                {
                    sqlcon.Open();
                    String var = "SP_DeletePP_FILE_V3";
                    //MessageBox.Show(var);
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = sqlcon;
                    cmd.CommandText = var;
                    cmd.ExecuteNonQuery();
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
                conString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR=YES'";
                conString = string.Format(conString, excelPath);
                using (OleDbConnection excel_con = new OleDbConnection(conString))
                {
                    excel_con.Open();
                    string sheet1 = excel_con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null).Rows[0]["TABLE_NAME"].ToString();
                    DataTable dtExcelData = new DataTable();

                    //[OPTIONAL]: It is recommended as otherwise the data will be considered as String by default.
                    dtExcelData.Columns.AddRange(new DataColumn[10] { new DataColumn("PP-ITEM-NO", typeof(int)),
                new DataColumn("PP-DATE", typeof(DateTime)),
                new DataColumn("PP-OP-IND", typeof(decimal)),
                new DataColumn("PP-CL-IND", typeof(decimal)),
                new DataColumn("PP-S1C1", typeof(decimal)),
                new DataColumn("PP-S1C2", typeof(decimal)),
                new DataColumn("PP-S1C3", typeof(decimal)),
                new DataColumn("HI VOL", typeof(decimal)),
                new DataColumn("PP-IRES", typeof(string)),
                new DataColumn("PP-ARES", typeof(decimal))
                });

                    using (OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [" + sheet1 + "]", excel_con))
                    {
                        oda.Fill(dtExcelData);
                    }
                    excel_con.Close();

                    using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(sqlcon))
                    {
                        //Set the database table name
                        sqlBulkCopy.DestinationTableName = "dbo.[PP-FILE]";

                        //[OPTIONAL]: Map the Excel columns with that of the database table
                        sqlBulkCopy.ColumnMappings.Add("PP-ITEM-NO", "ITEM-NO");
                        sqlBulkCopy.ColumnMappings.Add("PP-DATE", "PP-DATE");
                        sqlBulkCopy.ColumnMappings.Add("PP-OP-IND", "PP-OP");
                        sqlBulkCopy.ColumnMappings.Add("PP-CL-IND", "PP-CL");
                        sqlBulkCopy.ColumnMappings.Add("PP-S1C1", "PP-C1");
                        sqlBulkCopy.ColumnMappings.Add("PP-S1C2", "PP-C2");
                        sqlBulkCopy.ColumnMappings.Add("PP-S1C3", "PP-C3");
                        sqlBulkCopy.ColumnMappings.Add("HI VOL", "HI VOL");
                        sqlBulkCopy.ColumnMappings.Add("PP-IRES", "PP-IRES");
                        sqlBulkCopy.ColumnMappings.Add("PP-ARES", "PP-ARES");
                        sqlcon.Open();
                        sqlBulkCopy.WriteToServer(dtExcelData);
                        MessageBox.Show("PP File Uploaded successfully.");
                        sqlcon.Close();
                    }
                }
            }
            else if (excelPath.Contains("PR-FILE-WITH DATE - V2.xlsx"))
            {
                try
                {
                    sqlcon.Open();
                    String var = "SP_DeletePR_FILE_V3";
                    //MessageBox.Show(var);
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = sqlcon;
                    cmd.CommandText = var;
                    cmd.ExecuteNonQuery();
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
                conString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR=YES'";
                conString = string.Format(conString, excelPath);
                using (OleDbConnection excel_con = new OleDbConnection(conString))
                {
                    excel_con.Open();
                    string sheet1 = excel_con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null).Rows[0]["TABLE_NAME"].ToString();
                    DataTable dtExcelData = new DataTable();

                    //[OPTIONAL]: It is recommended as otherwise the data will be considered as String by default.
                    dtExcelData.Columns.AddRange(new DataColumn[24] { new DataColumn("PR-ITEM-NO", typeof(int)),
                new DataColumn("PR-DATE", typeof(DateTime)),
                new DataColumn("PROP", typeof(decimal)),
                new DataColumn("PR-S2C1", typeof(decimal)),
                new DataColumn("PR-S2-C2", typeof(decimal)),
                new DataColumn("PR-S2C3", typeof(decimal)),
                new DataColumn("PR-N1", typeof(decimal)),
                new DataColumn("PR-N2", typeof(decimal)),
                new DataColumn("PR-N3", typeof(decimal)),
                new DataColumn("PR-N4", typeof(decimal)),
                new DataColumn("PR-N5", typeof(decimal)),
                new DataColumn("PR-N6", typeof(decimal)),
                new DataColumn("PR-N7", typeof(decimal)),
                new DataColumn("PR-N8", typeof(decimal)),
                new DataColumn("PR(N1-N2)", typeof(decimal)),
                new DataColumn("PR-AVG", typeof(decimal)),
                new DataColumn("PR-S1C1", typeof(decimal)),
                new DataColumn("PR-S1C2", typeof(decimal)),
                new DataColumn("PR-S1C3", typeof(decimal)),
                new DataColumn("PRCL-IND", typeof(string)),
                new DataColumn("PR-330IND", typeof(decimal)),
                new DataColumn("PRFIN-IND", typeof(decimal)),
                new DataColumn("PR-IRES", typeof(string)),
                new DataColumn("PR-ARES", typeof(string))
                
            });

                    using (OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [" + sheet1 + "]", excel_con))
                    {
                        oda.Fill(dtExcelData);
                    }
                    excel_con.Close();

                    using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(sqlcon))
                    {
                        //Set the database table name
                        sqlBulkCopy.DestinationTableName = "dbo.[PR-FILE-V3]";

                        //[OPTIONAL]: Map the Excel columns with that of the database table
                        sqlBulkCopy.ColumnMappings.Add("PR-ITEM-NO", "PR-ITEM-NO");
                        sqlBulkCopy.ColumnMappings.Add("PR-DATE", "PR-DATE");
                        sqlBulkCopy.ColumnMappings.Add("PROP", "PROP");
                        sqlBulkCopy.ColumnMappings.Add("PR-S2C1", "PR-S2C1");
                        sqlBulkCopy.ColumnMappings.Add("PR-S2-C2", "PR-S2-C2");
                        sqlBulkCopy.ColumnMappings.Add("PR-S2C3", "PR-S2C3");
                        sqlBulkCopy.ColumnMappings.Add("PR-N1", "PR-N1");
                        sqlBulkCopy.ColumnMappings.Add("PR-N2", "PR-N2");
                        sqlBulkCopy.ColumnMappings.Add("PR-N3", "PR-N3");
                        sqlBulkCopy.ColumnMappings.Add("PR-N4", "PR-N4");
                        sqlBulkCopy.ColumnMappings.Add("PR-N5", "PR-N5");
                        sqlBulkCopy.ColumnMappings.Add("PR-N6", "PR-N6");
                        sqlBulkCopy.ColumnMappings.Add("PR-N7", "PR-N7");
                        sqlBulkCopy.ColumnMappings.Add("PR-N8", "PR-N8");
                        sqlBulkCopy.ColumnMappings.Add("PR(N1-N2)", "PR(N1-N2)");
                        sqlBulkCopy.ColumnMappings.Add("PR-AVG", "PR-AVG");
                        sqlBulkCopy.ColumnMappings.Add("PR-S1C1", "PR-S1C1");
                        sqlBulkCopy.ColumnMappings.Add("PR-S1C2", "PR-S1C2");
                        sqlBulkCopy.ColumnMappings.Add("PR-S1C3", "PR-S1C3");
                        sqlBulkCopy.ColumnMappings.Add("PRCL-IND", "PRCL-IND");
                        sqlBulkCopy.ColumnMappings.Add("PR-330IND", "PR-330IND");
                        sqlBulkCopy.ColumnMappings.Add("PRFIN-IND", "PRFIN-IND");
                        sqlBulkCopy.ColumnMappings.Add("PR-IRES", "PR-IRES");
                        sqlBulkCopy.ColumnMappings.Add("PR-ARES", "PR-ARES");
                        
                        sqlcon.Open();
                        sqlBulkCopy.WriteToServer(dtExcelData);
                        MessageBox.Show("PR File Uploaded successfully.");
                        sqlcon.Close();
                    }
                }
                try
                {
                    sqlcon.Open();
                    String var = "SP_PopulateProjectHealthEntry";
                    //MessageBox.Show(var);
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = sqlcon;
                    cmd.CommandText = var;
                    cmd.ExecuteNonQuery();
                }
                catch(Exception ex)
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

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

            try
            {
                sqlcon.Open();
                String var = "SP_P2LoadFinalHealthData";
                //MessageBox.Show(var);
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = sqlcon;
                cmd.CommandText = var;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@NoOfCells", SqlDbType.VarChar).Value = txtEnternoofcells.Text;
                //cmd.Parameters.Add("@ITEM_NO", SqlDbType.VarChar).Value = txtItemNoP2.Text;
                cmd.CommandTimeout = 3600000;
                cmd.ExecuteNonQuery();
                MessageBox.Show("Calculation Completed Successfully.");
                SqlDataAdapter sda = new SqlDataAdapter("SELECT [ITEM_NO],	[PR-DATE],	[PRCL-IND],	[330-IND],	[S1C2],	[UD1],	[PRCL/PRCL-IND-CNT1],	[UD2],	[PRCL/PRCL-IND-CNT2],	[UD3],	[PRCL/PRCL-IND-CNT3],	[PRCL/C2-CNT1],	[PRCL/C2-CNT2],	[PRCL/C3-CNT3],	[RED/INC],	[PP-RNG-UD],	[GAP-LMT],	[IND-GAP],	[LIND],	[MIND],	[Res9],	[Res8],	[Res7],	[Res6],	[Res5],	[Res4],	[Res3],	[Res2],	[Res1],	[TOT], [ALL-CNT(P/N)],	[GREEN],[ACC/TOT],[3COL/NET], [ACC/3COL],	[COMM],	[IRES], [PRFIN-IND],[UD1] AS [_UD1],[PRCL/330IND-R1-C],[UD2] AS [_UD2],[PRCL/330IND-R2-C],[UD3] AS [_UD3],[PRCL/330IND-R3-C], [FLAG],[PR-S1C1],[PR-S1C2],[PR-S1C3],[MinNoOfCells] FROM [dbo].[ProjectHelathEntryResultP2] ORDER BY ITEM_NO asc, (case when flag ='330IND->PPCL' then 3 else len(FLAG) end) asc,[RED/INC] desc", sqlcon);
                DataSet ds = new DataSet();
                sda.Fill(ds, "ProjectHelathEntryResult");
                dgv3.DataSource = ds;
                dgv3.DataMember = "ProjectHelathEntryResult";

                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                using (ExcelPackage excelPackage = new ExcelPackage())
                {

                    //Set some properties of the Excel document
                    excelPackage.Workbook.Properties.Author = "Sayan";
                    excelPackage.Workbook.Properties.Title = "P2-INDIVIDUAL-DETAIL-OUTPUTFILE-V1";
                    excelPackage.Workbook.Properties.Subject = "P2-INDIVIDUAL-DETAIL-OUTPUTFILE-V1";
                    excelPackage.Workbook.Properties.Created = DateTime.Now;

                    //Create the WorkSheet
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("P2-DETAIL-OUTPUT-FILE");

                    for (int i = 0; i <= dgv3.Columns.Count - 1; i++)
                    {
                        worksheet.Cells[1, i + 1].Value = dgv3.Columns[i].HeaderText;
                    }

                    /*And the information of your data*/
                    for (int i = 0; i <= dgv3.RowCount - 1; i++)
                    {
                        for (int j = 0; j <= dgv3.ColumnCount - 1; j++)
                        {
                            DataGridViewCell cell = dgv3[j, i];
                            worksheet.Cells[i + 2, j + 1].Value = cell.Value;

                        }
                     }

                        //Save your file
                        FileInfo fi = new FileInfo(@"C:\ProjectHealthApplication\Export\P2-INDIVIDUAL-DETAIL-OUTPUTFILE-V1.xlsx");
                        excelPackage.SaveAs(fi);
                    }
                    try
                    {
                        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                        ExcelPackage ExcelPkg = new ExcelPackage();

                        FileInfo fi = new FileInfo(@"C:\ProjectHealthApplication\Export\P2-INDIVIDUAL-DETAIL-OUTPUTFILE-V1.xlsx");
                        using (ExcelPackage excelPackage = new ExcelPackage(fi))
                        {
                            //Get a WorkSheet by name. If the worksheet doesn't exist, throw an exeption
                            ExcelWorksheet namedWorksheet = excelPackage.Workbook.Worksheets["P2-DETAIL-OUTPUT-FILE"];

                            //If you don't know if a worksheet exists, you could use LINQ,
                            //So it doesn't throw an exception, but return null in case it doesn't find it
                            ExcelWorksheet anotherWorksheet =
                                excelPackage.Workbook.Worksheets.FirstOrDefault(x => x.Name == "P2-DETAIL-OUTPUT-FILE");


                            //ExcelRange Rng1= namedWorksheet.Cells[2,1, dgv3.RowCount,2]
                            //int j = 0;
                            for (int k = 0; k < dgv3.ColumnCount; k++)
                            {

                                using (ExcelRange Rng = namedWorksheet.Cells[1, 1, 1, k + 1])
                                {
                                    namedWorksheet.Row(1).Height = 40;
                                    namedWorksheet.Column(k + 1).Width = 6;
                                    namedWorksheet.Column(k + 1).Style.WrapText = true;
                                    namedWorksheet.Cells[1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                                    //namedWorksheet.Cells[1, 2].Style.
                                    namedWorksheet.Cells[1, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[1, 2].Style.Fill.BackgroundColor.SetColor(Color.Yellow);

                                    namedWorksheet.Cells[1, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[1, 3].Style.Fill.BackgroundColor.SetColor(Color.Gold);

                                    namedWorksheet.Cells[1, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[1, 4].Style.Fill.BackgroundColor.SetColor(Color.Yellow);

                                    //namedWorksheet.Cells[1, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    //namedWorksheet.Cells[1, 6].Style.Fill.BackgroundColor.SetColor(Color.DeepSkyBlue);

                                    namedWorksheet.Cells[1, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[1, 6].Style.Fill.BackgroundColor.SetColor(Color.DeepSkyBlue);

                                    namedWorksheet.Cells[1, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[1, 7].Style.Fill.BackgroundColor.SetColor(Color.LightCyan);

                                    namedWorksheet.Cells[1, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[1, 8].Style.Fill.BackgroundColor.SetColor(Color.DeepSkyBlue);

                                    namedWorksheet.Cells[1, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[1, 9].Style.Fill.BackgroundColor.SetColor(Color.LightCyan);

                                    namedWorksheet.Cells[1, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[1, 10].Style.Fill.BackgroundColor.SetColor(Color.DeepSkyBlue);

                                    namedWorksheet.Cells[1, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[1, 11].Style.Fill.BackgroundColor.SetColor(Color.LightCyan);

                                    namedWorksheet.Cells[1, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[1, 12].Style.Fill.BackgroundColor.SetColor(Color.DarkGray);
                                    namedWorksheet.Cells[1, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[1, 13].Style.Fill.BackgroundColor.SetColor(Color.DarkGray);
                                    namedWorksheet.Cells[1, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[1, 14].Style.Fill.BackgroundColor.SetColor(Color.DarkGray);


                                    namedWorksheet.Cells[1, 15].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[1, 15].Style.Fill.BackgroundColor.SetColor(Color.Violet);

                                    namedWorksheet.Cells[1, 16].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[1, 16].Style.Fill.BackgroundColor.SetColor(Color.PeachPuff);

                                    namedWorksheet.Cells[1, 19].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[1, 19].Style.Fill.BackgroundColor.SetColor(Color.LightSkyBlue);
                                    namedWorksheet.Cells[1, 20].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[1, 20].Style.Fill.BackgroundColor.SetColor(Color.LightSkyBlue);

                                    namedWorksheet.Cells[1, 30].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[1, 30].Style.Fill.BackgroundColor.SetColor(Color.Green);
                                    namedWorksheet.Cells[1, 31].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[1, 31].Style.Fill.BackgroundColor.SetColor(Color.Green);
                                    //namedWorksheet.Cells[1, 32].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    //namedWorksheet.Cells[1, 32].Style.Fill.BackgroundColor.SetColor(Color.Green);

                                    namedWorksheet.Cells[1, 34].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[1, 34].Style.Fill.BackgroundColor.SetColor(Color.Green);
                                    namedWorksheet.Cells[1, 37].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[1, 37].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                                    namedWorksheet.Cells[1, 39].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[1, 39].Style.Fill.BackgroundColor.SetColor(Color.LightGoldenrodYellow);
                                    namedWorksheet.Cells[1, 40].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[1, 40].Style.Fill.BackgroundColor.SetColor(Color.LightGoldenrodYellow);
                                    namedWorksheet.Cells[1, 41].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[1, 41].Style.Fill.BackgroundColor.SetColor(Color.LightGoldenrodYellow);
                                    namedWorksheet.Cells[1, 42].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[1, 42].Style.Fill.BackgroundColor.SetColor(Color.LightGoldenrodYellow);
                                    namedWorksheet.Cells[1, 43].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[1, 43].Style.Fill.BackgroundColor.SetColor(Color.LightGoldenrodYellow);
                                    namedWorksheet.Cells[1, 44].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    namedWorksheet.Cells[1, 44].Style.Fill.BackgroundColor.SetColor(Color.LightGoldenrodYellow);

                                    namedWorksheet.View.FreezePanes(2, 1);
                                    //Rng.AutoFitColumns();

                                    Rng.Style.Font.Size = 8;
                                    Rng.Style.Font.Bold = true;
                                    Rng.Style.Font.Color.SetColor(Color.Red);
                                    Rng.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                    Rng.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                    Rng.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                    Rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                    Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    Rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                }

                                using (ExcelRange Rng = namedWorksheet.Cells[2, (k + 1), dgv3.RowCount, (k + 1)])
                                {

                                    if (k == 2)
                                    {
                                        Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        Rng.Style.Fill.BackgroundColor.SetColor(Color.Gold);

                                    }
                                    else if (k == 3)
                                    {
                                        Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        Rng.Style.Fill.BackgroundColor.SetColor(Color.Green);
                                    }
                                    else if (k == 5 || k == 7 || k == 9)
                                    {
                                        Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        Rng.Style.Fill.BackgroundColor.SetColor(Color.DeepSkyBlue);
                                    }
                                    else if (k == 6 || k == 8 || k == 10)
                                    {
                                        Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        Rng.Style.Fill.BackgroundColor.SetColor(Color.LightCyan);
                                    }
                                    else if (k == 11 || k == 12 || k == 13)
                                    {
                                        Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        Rng.Style.Fill.BackgroundColor.SetColor(Color.DarkGray);
                                    }
                                    else if (k == 14)
                                    {
                                        Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        Rng.Style.Fill.BackgroundColor.SetColor(Color.Violet);

                                    }
                                    else if (k == 15)
                                    {
                                        Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        Rng.Style.Fill.BackgroundColor.SetColor(Color.PeachPuff);
                                    }
                                    else if (k == 18 || k == 19)
                                    {
                                        Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        Rng.Style.Fill.BackgroundColor.SetColor(Color.LightSkyBlue);
                                    }
                                    else if (k >= 20 && k <= 28)
                                    {
                                        Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        Rng.Style.Fill.BackgroundColor.SetColor(Color.Yellow);

                                    }
                                    else if (k >= 29 && k <= 30)
                                    {
                                        Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        Rng.Style.Fill.BackgroundColor.SetColor(Color.Green);
                                    }
                                    else if (k == 33)
                                    {
                                        Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        Rng.Style.Fill.BackgroundColor.SetColor(Color.Green);
                                    }
                                    else if (k == 36)
                                    {
                                        Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        Rng.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                                    }
                                    else if (k == 38 || k == 40 || k == 42)
                                    {
                                        Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        Rng.Style.Fill.BackgroundColor.SetColor(Color.LawnGreen);
                                    }
                                    else
                                    {
                                        Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        Rng.Style.Fill.BackgroundColor.SetColor(Color.White);
                                    }
                                    namedWorksheet.Column(k + 1).Width = 6;
                                    namedWorksheet.Column(k + 1).Style.WrapText = true;
                                    Rng.Style.Font.Size = 8;
                                    Rng.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                    Rng.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                    Rng.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                    Rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                    Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    Rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                    //Rng.AutoFitColumns();

                                }


                            }
                            //RED INC COLORING

                            for (int i = 2; i <= dgv3.RowCount; i++)
                            {
                                if (namedWorksheet.Cells[i, 15].Value.ToString() == "RED")
                                {
                                    if (Convert.ToDecimal(namedWorksheet.Cells[i, 4].Value.ToString()) < Convert.ToDecimal(namedWorksheet.Cells[i, 3].Value.ToString()))
                                    {
                                        namedWorksheet.Cells[i, 15].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        namedWorksheet.Cells[i, 15].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                        namedWorksheet.Cells[i, 16].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        namedWorksheet.Cells[i, 16].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                        namedWorksheet.Cells[i, 17].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        namedWorksheet.Cells[i, 17].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                        namedWorksheet.Cells[i, 18].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        namedWorksheet.Cells[i, 18].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                        //namedWorksheet.Cells[i, 19].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        //namedWorksheet.Cells[i, 19].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                        //namedWorksheet.Cells[i, 20].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        //namedWorksheet.Cells[i, 20].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                        namedWorksheet.Cells[i, 21].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        namedWorksheet.Cells[i, 21].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                        namedWorksheet.Cells[i, 22].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        namedWorksheet.Cells[i, 22].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                        namedWorksheet.Cells[i, 23].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        namedWorksheet.Cells[i, 23].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                        namedWorksheet.Cells[i, 24].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        namedWorksheet.Cells[i, 24].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                        namedWorksheet.Cells[i, 25].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        namedWorksheet.Cells[i, 25].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                        namedWorksheet.Cells[i, 26].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        namedWorksheet.Cells[i, 26].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                        namedWorksheet.Cells[i, 27].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        namedWorksheet.Cells[i, 27].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                        namedWorksheet.Cells[i, 28].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        namedWorksheet.Cells[i, 28].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                        namedWorksheet.Cells[i, 29].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        namedWorksheet.Cells[i, 29].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                        //namedWorksheet.Cells[i, 30].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        //namedWorksheet.Cells[i, 30].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                        //namedWorksheet.Cells[i, 31].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        //namedWorksheet.Cells[i, 31].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                        namedWorksheet.Cells[i, 32].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        namedWorksheet.Cells[i, 32].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                        namedWorksheet.Cells[i, 33].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        namedWorksheet.Cells[i, 33].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                        //namedWorksheet.Cells[i, 34].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        //namedWorksheet.Cells[i, 34].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                        namedWorksheet.Cells[i, 35].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        namedWorksheet.Cells[i, 35].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);

                                    }
                                }
                                if (namedWorksheet.Cells[i, 15].Value.ToString() == "INC")
                                {
                                    if (Convert.ToDecimal(namedWorksheet.Cells[i, 4].Value.ToString()) > Convert.ToDecimal(namedWorksheet.Cells[i, 3].Value.ToString()))
                                    {
                                        //MessageBox.Show("INC Hi...");
                                        namedWorksheet.Cells[i, 15].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        namedWorksheet.Cells[i, 15].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                        namedWorksheet.Cells[i, 16].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        namedWorksheet.Cells[i, 16].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                        namedWorksheet.Cells[i, 17].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        namedWorksheet.Cells[i, 17].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                        namedWorksheet.Cells[i, 18].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        namedWorksheet.Cells[i, 18].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                        //namedWorksheet.Cells[i, 19].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        //namedWorksheet.Cells[i, 19].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                        //namedWorksheet.Cells[i, 20].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        //namedWorksheet.Cells[i, 20].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                        namedWorksheet.Cells[i, 21].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        namedWorksheet.Cells[i, 21].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                        namedWorksheet.Cells[i, 22].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        namedWorksheet.Cells[i, 22].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                        namedWorksheet.Cells[i, 23].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        namedWorksheet.Cells[i, 23].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                        namedWorksheet.Cells[i, 24].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        namedWorksheet.Cells[i, 24].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                        namedWorksheet.Cells[i, 25].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        namedWorksheet.Cells[i, 25].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                        namedWorksheet.Cells[i, 26].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        namedWorksheet.Cells[i, 26].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                        namedWorksheet.Cells[i, 27].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        namedWorksheet.Cells[i, 27].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                        namedWorksheet.Cells[i, 28].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        namedWorksheet.Cells[i, 28].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                        namedWorksheet.Cells[i, 29].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        namedWorksheet.Cells[i, 29].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                        //namedWorksheet.Cells[i, 30].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        //namedWorksheet.Cells[i, 30].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                        //namedWorksheet.Cells[i, 31].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        //namedWorksheet.Cells[i, 31].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                        namedWorksheet.Cells[i, 32].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        namedWorksheet.Cells[i, 32].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                        namedWorksheet.Cells[i, 33].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        namedWorksheet.Cells[i, 33].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                        //namedWorksheet.Cells[i, 34].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        //namedWorksheet.Cells[i, 34].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                        namedWorksheet.Cells[i, 35].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        namedWorksheet.Cells[i, 35].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);
                                    }
                                }
                                //MessageBox.Show(namedWorksheet.Cells[i, 1].Value.ToString());
                                //MessageBox.Show(namedWorksheet.Cells[i + 1, 1].Value.ToString());
                                if (i < dgv3.RowCount)
                                {
                                    int j = Convert.ToInt32(namedWorksheet.Cells[i, 1].Value.ToString());
                                    int k = Convert.ToInt32(namedWorksheet.Cells[i + 1, 1].Value.ToString());

                                    if (j != k)
                                    {
                                        //MessageBox.Show(dgv3.ColumnCount.ToString());
                                        for (int l = 1; l <= dgv3.ColumnCount; l++)
                                        {
                                            //MessageBox.Show(namedWorksheet.Cells[i, l].Value.ToString());
                                            //namedWorksheet.Cells[i, l].Style.Border.Top.Style = ExcelBorderStyle.Thick;
                                            //namedWorksheet.Cells[i, l].Style.Border.Left.Style = ExcelBorderStyle.Thick;
                                            //namedWorksheet.Cells[i, l].Style.Border.Right.Style = ExcelBorderStyle.Thick;
                                            //namedWorksheet.InsertRow(i, 1);
                                            //namedWorksheet.Cells[i, l].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
                                        }
                                        //MessageBox.Show(i.ToString());
                                        //namedWorksheet.InsertRow(i, 1);
                                    }


                                }


                            }
                        //END OF RED INC COLORING
                        using (ExcelRange Rng = namedWorksheet.Cells[1, 1, 1, dgv3.ColumnCount])
                            {

                                Rng.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                Rng.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                Rng.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                Rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                            }
                            
                            int cntRow = dgv3.RowCount;
                            Decimal P = 0;
                            Decimal Q = 0;
                            Decimal R = 0;
                            Decimal S = 0;
                            Decimal Q1 = 0;
                            Decimal S1 = 0;
                            String v11 = "[0/0]";
                            String t = "[0/0]";
                            String v = "[0/0]";
                            for (int s = 2; s < cntRow; s++)
                            {
                                //MessageBox.Show(namedWorksheet.Cells[s, 45].Value.ToString());
                                String j1 = (String.IsNullOrEmpty(namedWorksheet.Cells[s, 45].Value.ToString()) ? "dummy" : namedWorksheet.Cells[s, 45].Value.ToString().Trim());
                                String k1 = (String.IsNullOrEmpty(namedWorksheet.Cells[s + 1, 45].Value.ToString()) ? "dummy" : namedWorksheet.Cells[s + 1, 45].Value.ToString().Trim());
                                //MessageBox.Show(j1.ToString());
                                //MessageBox.Show(k1.ToString());
                                
                                if (j1 == "330IND->PPCL" && k1 == "PRCL->PPCL")
                                {
                                    namedWorksheet.InsertRow(s + 1, 1);
                                    namedWorksheet.Cells[s + 1, 15].Value = "TOTAL_330IND->PPCL";
                                namedWorksheet.Cells[s + 1, 45].Value = "TOTAL_330IND->PPCL";
                                namedWorksheet.Cells[s + 1, 30].Value = Convert.ToDecimal(String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[s, 30].Value)) ? "0" : Convert.ToString(namedWorksheet.Cells[s, 30].Value))
                                        + Convert.ToDecimal(String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[s - 1, 30].Value)) ? "0" : Convert.ToString(namedWorksheet.Cells[s - 1, 30].Value));
                                //+ Convert.ToDecimal(String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[s - 2, 30].Value)) ? "0" : Convert.ToString(namedWorksheet.Cells[s - 2, 30].Value));


                                namedWorksheet.Cells[s + 1, 34].Value = Convert.ToDecimal(String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[s, 34].Value)) ? "0" : Convert.ToString(namedWorksheet.Cells[s, 34].Value))
                                        + Convert.ToDecimal(String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[s - 1, 34].Value)) ? "0" : Convert.ToString(namedWorksheet.Cells[s - 1, 34].Value));
                                            //+ Convert.ToDecimal(String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[s - 2, 34].Value)) ? "0" : Convert.ToString(namedWorksheet.Cells[s - 2, 34].Value));

                                    var t1 = String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[s, 31].Value)) ? "[0/0]" : Convert.ToString(namedWorksheet.Cells[s, 31].Value);
                                    var u1 = String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[s - 1, 31].Value)) ? "[0/0]" : Convert.ToString(namedWorksheet.Cells[s - 1, 31].Value);
                                    //var v1 = String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[s - 2, 31].Value)) ? "[0/0]" : Convert.ToString(namedWorksheet.Cells[s - 2, 31].Value);
                                    var v1 = "[0/0]";
                                    var Str = "SELECT dbo.fn_CalculateCountSum(" + "'" + t1.ToString().Trim() + "','" + u1.ToString().Trim() + "','" + v1.ToString().Trim() + "'" + ")";

                                    using (SqlCommand cmd1 = new SqlCommand(Str, sqlcon))
                                    {
                                        var result = cmd1.ExecuteScalar();
                                        namedWorksheet.Cells[s + 1, 31].Value = result;
                                    }

                                    cntRow++;
                                    //P = Convert.ToDecimal(namedWorksheet.Cells[s + 1, 30].Value);
                                }
                            if (j1 == "PRCL->PPCL" && k1 == "PRCL/330IND->PPOP/PPCL")
                            {
                                namedWorksheet.InsertRow(s + 1, 1);
                                namedWorksheet.Cells[s + 1, 15].Value = "TOTAL_PRCL->PPCL";
                                namedWorksheet.Cells[s + 1, 45].Value = "TOTAL_PRCL->PPCL";
                                namedWorksheet.Cells[s + 1, 30].Value = Convert.ToDecimal(String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[s, 30].Value)) ? "0" : Convert.ToString(namedWorksheet.Cells[s, 30].Value))
                                        + Convert.ToDecimal(String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[s - 1, 30].Value)) ? "0" : Convert.ToString(namedWorksheet.Cells[s - 1, 30].Value));
                                //+ Convert.ToDecimal(String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[s - 2, 30].Value)) ? "0" : Convert.ToString(namedWorksheet.Cells[s - 2, 30].Value));


                                namedWorksheet.Cells[s + 1, 34].Value = Convert.ToDecimal(String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[s, 34].Value)) ? "0" : Convert.ToString(namedWorksheet.Cells[s, 34].Value))
                                        + Convert.ToDecimal(String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[s - 1, 34].Value)) ? "0" : Convert.ToString(namedWorksheet.Cells[s - 1, 34].Value));
                                //+ Convert.ToDecimal(String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[s - 2, 34].Value)) ? "0" : Convert.ToString(namedWorksheet.Cells[s - 2, 34].Value));

                                var t1 = String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[s, 31].Value)) ? "[0/0]" : Convert.ToString(namedWorksheet.Cells[s, 31].Value);
                                var u1 = String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[s - 1, 31].Value)) ? "[0/0]" : Convert.ToString(namedWorksheet.Cells[s - 1, 31].Value);
                                //var v1 = String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[s - 2, 31].Value)) ? "[0/0]" : Convert.ToString(namedWorksheet.Cells[s - 2, 31].Value);
                                var v1 = "[0/0]";
                                var Str = "SELECT dbo.fn_CalculateCountSum(" + "'" + t1.ToString().Trim() + "','" + u1.ToString().Trim() + "','" + v1.ToString().Trim() + "'" + ")";

                                using (SqlCommand cmd1 = new SqlCommand(Str, sqlcon))
                                {
                                    var result = cmd1.ExecuteScalar();
                                    namedWorksheet.Cells[s + 1, 31].Value = result;
                                }

                                cntRow++;
                                //P = Convert.ToDecimal(namedWorksheet.Cells[s + 1, 30].Value);
                            }
                            if (j1 == "PRCL/330IND->PPOP/PPCL" && k1 == "330IND->PPCL")
                                {
                                    namedWorksheet.InsertRow(s + 1, 2);
                                    namedWorksheet.Cells[s + 1, 15].Value = "TOTAL_PRCL/330IND->PPOP/PPCL";
                                    namedWorksheet.Cells[s + 1, 45].Value = "TOTAL_PRCL/330IND->PPOP/PPCL";
                                    namedWorksheet.Cells[s + 2, 15].Value = "T_RI";
                                    namedWorksheet.Cells[s + 2, 45].Value = "T_RI";
                                namedWorksheet.Cells[s + 1, 30].Value = Convert.ToDecimal(String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[s, 30].Value)) ? "0" : Convert.ToString(namedWorksheet.Cells[s, 30].Value))
                                        + Convert.ToDecimal(String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[s - 1, 30].Value)) ? "0" : Convert.ToString(namedWorksheet.Cells[s - 1, 30].Value));
                                //+ Convert.ToDecimal(String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[s - 2, 30].Value)) ? "0" : Convert.ToString(namedWorksheet.Cells[s - 2, 30].Value));

                                namedWorksheet.Cells[s + 1, 34].Value = Convert.ToDecimal(String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[s, 34].Value)) ? "0" : Convert.ToString(namedWorksheet.Cells[s, 34].Value))
                                        + Convert.ToDecimal(String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[s - 1, 34].Value)) ? "0" : Convert.ToString(namedWorksheet.Cells[s - 1, 34].Value));
                                            //+ Convert.ToDecimal(String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[s - 2, 34].Value)) ? "0" : Convert.ToString(namedWorksheet.Cells[s - 2, 34].Value));

                                    var t2 = String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[s, 31].Value)) ? "[0/0]" : Convert.ToString(namedWorksheet.Cells[s, 31].Value);
                                    var u2 = String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[s - 1, 31].Value)) ? "[0/0]" : Convert.ToString(namedWorksheet.Cells[s - 1, 31].Value);
                                    //var v2 = String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[s - 2, 31].Value)) ? "[0/0]" : Convert.ToString(namedWorksheet.Cells[s - 2, 31].Value);
                                    var v2 = "[0/0]";

                                    var Str1 = "SELECT dbo.fn_CalculateCountSum(" + "'" + t2.ToString().Trim() + "','" + u2.ToString().Trim() + "','" + v2.ToString().Trim() + "'" + ")";

                                    using (SqlCommand cmd2 = new SqlCommand(Str1, sqlcon))
                                    {

                                        var result = cmd2.ExecuteScalar();
                                        namedWorksheet.Cells[s + 1, 31].Value = result;

                                    }
                                    //cntRow++;
                                    cntRow = cntRow + 2;
                                    //Q = Convert.ToDecimal(namedWorksheet.Cells[s + 1, 30].Value);
                                    //namedWorksheet.Cells[s + 2, 30].Value = (P + Q);
                                }

                            }
                            //MessageBox.Show(cntRow.ToString());
                            namedWorksheet.InsertRow(cntRow + 1, 2);
                            namedWorksheet.Cells[cntRow + 1, 15].Value = "TOTAL_PRCL/330IND->PPOP/PPCL";
                            namedWorksheet.Cells[cntRow + 1, 45].Value = "TOTAL_PRCL/330IND->PPOP/PPCL";
                            namedWorksheet.Cells[cntRow + 2, 15].Value = "T_RI";
                            namedWorksheet.Cells[cntRow + 2, 45].Value = "T_RI";
                        namedWorksheet.Cells[cntRow + 1, 30].Value = Convert.ToDecimal(String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[cntRow, 30].Value)) ? "0" : Convert.ToString(namedWorksheet.Cells[cntRow, 30].Value))
                                        + Convert.ToDecimal(String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[cntRow - 1, 30].Value)) ? "0" : Convert.ToString(namedWorksheet.Cells[cntRow - 1, 30].Value));
                        //+ Convert.ToDecimal(String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[cntRow - 2, 30].Value)) ? "0" : Convert.ToString(namedWorksheet.Cells[cntRow - 2, 30].Value));

                        namedWorksheet.Cells[cntRow + 1, 34].Value = Convert.ToDecimal(String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[cntRow, 34].Value)) ? "0" : Convert.ToString(namedWorksheet.Cells[cntRow, 34].Value))
                                        + Convert.ToDecimal(String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[cntRow - 1, 34].Value)) ? "0" : Convert.ToString(namedWorksheet.Cells[cntRow - 1, 34].Value));
                                            //+ Convert.ToDecimal(String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[cntRow - 2, 34].Value)) ? "0" : Convert.ToString(namedWorksheet.Cells[cntRow - 2, 34].Value));


                            var t3 = String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[cntRow, 31].Value)) ? "[0/0]" : Convert.ToString(namedWorksheet.Cells[cntRow, 31].Value);
                            var u3 = String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[cntRow - 1, 31].Value)) ? "[0/0]" : Convert.ToString(namedWorksheet.Cells[cntRow - 1, 31].Value);
                            //var v3 = String.IsNullOrEmpty(Convert.ToString(namedWorksheet.Cells[cntRow - 2, 31].Value)) ? "[0/0]" : Convert.ToString(namedWorksheet.Cells[cntRow - 2, 31].Value);
                            var v3 = "[0/0]";

                            var Str3 = "SELECT dbo.fn_CalculateCountSum(" + "'" + t3.ToString().Trim() + "','" + u3.ToString().Trim() + "','" + v3.ToString().Trim() + "'" + ")";

                            using (SqlCommand cmd3 = new SqlCommand(Str3, sqlcon))
                            {
                                var result = cmd3.ExecuteScalar();
                                namedWorksheet.Cells[cntRow + 1, 31].Value = result;
                            }

                            for (int mm = 2; mm <= cntRow + 2; mm++)
                            {
                                if (Convert.ToString(namedWorksheet.Cells[mm, 15].Value) == "TOTAL_330IND->PPCL")
                                {
                                    P = Convert.ToDecimal(namedWorksheet.Cells[mm, 30].Value);
                                    R = Convert.ToDecimal(namedWorksheet.Cells[mm, 34].Value);
                                    t = namedWorksheet.Cells[mm, 31].Value.ToString();
                                }
                                if (Convert.ToString(namedWorksheet.Cells[mm, 15].Value) == "TOTAL_PRCL->PPCL")
                                {
                                    Q = Convert.ToDecimal(namedWorksheet.Cells[mm, 30].Value);
                                    S = Convert.ToDecimal(namedWorksheet.Cells[mm, 34].Value);
                                    v = namedWorksheet.Cells[mm, 31].Value.ToString();
                                }
                                if (Convert.ToString(namedWorksheet.Cells[mm, 15].Value) == "TOTAL_PRCL/330IND->PPOP/PPCL")
                                {
                                    Q1 = Convert.ToDecimal(namedWorksheet.Cells[mm, 30].Value);
                                    S1 = Convert.ToDecimal(namedWorksheet.Cells[mm, 34].Value);
                                    v11 = namedWorksheet.Cells[mm, 31].Value.ToString();
                                }
                                if (Convert.ToString(namedWorksheet.Cells[mm, 15].Value) == "T_RI")
                                {
                                    for (int l = 1; l <= dgv3.ColumnCount; l++)
                                    {
                                        //namedWorksheet.Cells[mm, l].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        //namedWorksheet.Cells[mm, l].Style.Fill.BackgroundColor.SetColor(Color.Red);

                                        namedWorksheet.Cells[mm, l].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
                                        namedWorksheet.Cells[mm, l].Style.Border.Bottom.Color.SetColor(Color.Red);
                                    }
                                    namedWorksheet.Cells[mm, 30].Value = (P + Q + Q1);
                                    namedWorksheet.Cells[mm, 34].Value = (R + S + S1);

                                   // MessageBox.Show("t-->" + t.ToString() + "  v-->" + v.ToString() + "  v11-->" + v11.ToString());

                                    var Str4 = "SELECT dbo.fn_CalculateCountSum(" + "'" + t.ToString().Trim() + "','" + v.ToString().Trim() + "','" + v11.ToString().Trim() + "'" + ")";

                                    using (SqlCommand cmd4 = new SqlCommand(Str4, sqlcon))
                                    {
                                        var result = cmd4.ExecuteScalar();
                                        namedWorksheet.Cells[mm, 31].Value = result;
                                    }
                                }

                            }
                            for (int mn = 2; mn <= cntRow + 2; mn++)
                            {
                                if (Convert.ToString(namedWorksheet.Cells[mn, 15].Value) == "T_RI")
                                {

                                    var T_Value = Convert.ToDecimal(namedWorksheet.Cells[mn, 30].Value.ToString());
                                    var T_3COLValue = Convert.ToDecimal(namedWorksheet.Cells[mn, 34].Value.ToString());
                                    var T_CNTPN = namedWorksheet.Cells[mn, 31].Value.ToString();
                                    //MessageBox.Show(namedWorksheet.Cells[mn, 31].Value.ToString());
                                    //MessageBox.Show(namedWorksheet.Cells[mn - 1, 37].Value.ToString());
                                    namedWorksheet.Cells[mn, 37].Value = Convert.ToDecimal(String.IsNullOrEmpty(namedWorksheet.Cells[mn - 2, 37].Value.ToString()) ? "0" : namedWorksheet.Cells[mn - 2, 37].Value.ToString());
                                    namedWorksheet.Cells[mn, 38].Value = Convert.ToDecimal(String.IsNullOrEmpty(namedWorksheet.Cells[mn - 2, 38].Value.ToString()) ? "0" : namedWorksheet.Cells[mn - 2, 38].Value.ToString());

                                    var T_IRES = Convert.ToDecimal(String.IsNullOrEmpty(namedWorksheet.Cells[mn - 2, 37].Value.ToString()) ? "0" : namedWorksheet.Cells[mn - 2, 37].Value.ToString());

                                    var Str5 = "SELECT dbo.fn_FinalCountPN(" + "'" + T_CNTPN.ToString().Trim() + "'" + ")";

                                    using (SqlCommand cmd5 = new SqlCommand(Str5, sqlcon))
                                    {
                                        var result5 = cmd5.ExecuteScalar();
                                    //namedWorksheet.Cells[cntRow + 1, 31].Value = result;

                                    if (T_Value > 0)
                                    {
                                        if (T_IRES >= 0)
                                        {
                                            namedWorksheet.Cells[mn, 33].Value = 'C';
                                        }
                                        else if (T_IRES < 0)
                                        {
                                            namedWorksheet.Cells[mn, 33].Value = 'W';
                                        }
                                    }
                                    else if (T_Value==0)
                                    {
                                        if (T_IRES >= 0)
                                        {
                                            namedWorksheet.Cells[mn, 33].Value = 'W';
                                        }
                                        else if (T_IRES < 0)
                                        {
                                            namedWorksheet.Cells[mn, 33].Value = 'C';
                                        }
                                    }
                                    else if (T_Value < 0)
                                    {
                                        if (T_IRES >= 0)
                                        {
                                            namedWorksheet.Cells[mn, 33].Value = 'W';
                                        }
                                        else if (T_IRES < 0)
                                        {
                                            namedWorksheet.Cells[mn, 33].Value = 'C';
                                        }
                                    }
                                    /*
                                    if (T_Value > 0)
                                    {
                                        if (Convert.ToInt32(result5.ToString()) == 1)
                                        {
                                            if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'C';
                                            }
                                            else if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'W';
                                            }

                                        }
                                        if (Convert.ToInt32(result5.ToString()) == 0)
                                        {
                                            if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'W';
                                            }
                                            else if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'C';
                                            }

                                        }
                                        if (Convert.ToInt32(result5.ToString()) == 2)
                                        {
                                            if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'C';
                                            }
                                            else if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'W';
                                            }

                                        }
                                    }
                                    else if (T_Value < 0)
                                    {
                                        if (Convert.ToInt32(result5.ToString()) == 0)
                                        {
                                            if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'C';
                                            }
                                            else if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'W';
                                            }

                                        }
                                        if (Convert.ToInt32(result5.ToString()) == 1)
                                        {
                                            if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'W';
                                            }
                                            else if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'C';
                                            }

                                        }
                                        if (Convert.ToInt32(result5.ToString()) == 2)
                                        {
                                            if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'C';
                                            }
                                            else if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 33].Value = 'W';
                                            }

                                        }
                                    }
                                    */

                                    //3col

                                    if (T_3COLValue > 0)
                                    {
                                        if (T_IRES >= 0)
                                        {
                                            namedWorksheet.Cells[mn, 33].Value = 'C';
                                        }
                                        else if (T_IRES < 0)
                                        {
                                            namedWorksheet.Cells[mn, 33].Value = 'W';
                                        }
                                    }
                                    else if (T_3COLValue == 0)
                                    {
                                        if (T_IRES >= 0)
                                        {
                                            namedWorksheet.Cells[mn, 33].Value = 'W';
                                        }
                                        else if (T_IRES < 0)
                                        {
                                            namedWorksheet.Cells[mn, 33].Value = 'C';
                                        }
                                    }
                                    else if (T_3COLValue < 0)
                                    {
                                        if (T_IRES >= 0)
                                        {
                                            namedWorksheet.Cells[mn, 33].Value = 'W';
                                        }
                                        else if (T_IRES < 0)
                                        {
                                            namedWorksheet.Cells[mn, 33].Value = 'C';
                                        }
                                    }
                                    /*
                                    if (T_3COLValue > 0)
                                    {
                                        if (Convert.ToInt32(result5.ToString()) == 1)
                                        {
                                            if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'C';
                                            }
                                            else if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'W';
                                            }

                                        }
                                        if (Convert.ToInt32(result5.ToString()) == 0)
                                        {
                                            if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'W';
                                            }
                                            else if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'C';
                                            }

                                        }
                                        if (Convert.ToInt32(result5.ToString()) == 2)
                                        {
                                            if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'C';
                                            }
                                            else if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'W';
                                            }

                                        }
                                    }
                                    else if (T_3COLValue < 0)
                                    {
                                        if (Convert.ToInt32(result5.ToString()) == 0)
                                        {
                                            if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'C';
                                            }
                                            else if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'W';
                                            }

                                        }
                                        if (Convert.ToInt32(result5.ToString()) == 1)
                                        {
                                            if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'W';
                                            }
                                            else if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'C';
                                            }

                                        }
                                        if (Convert.ToInt32(result5.ToString()) == 2)
                                        {
                                            if (T_IRES <= 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'C';
                                            }
                                            else if (T_IRES > 0)
                                            {
                                                namedWorksheet.Cells[mn, 35].Value = 'W';
                                            }

                                        }
                                    }

                                    */

                                }

                                }

                                if (Convert.ToString(namedWorksheet.Cells[mn, 15].Value) == "TOTAL_330IND->PPCL")
                                {
                                    var T_Value = Convert.ToDecimal(namedWorksheet.Cells[mn, 30].Value.ToString());
                                    var T_3COLValue = Convert.ToDecimal(namedWorksheet.Cells[mn, 34].Value.ToString());
                                    var T_CNTPN = namedWorksheet.Cells[mn, 31].Value.ToString();
                                    //MessageBox.Show(namedWorksheet.Cells[mn, 31].Value.ToString());
                                    //MessageBox.Show(namedWorksheet.Cells[mn - 1, 37].Value.ToString());
                                    namedWorksheet.Cells[mn, 37].Value = Convert.ToDecimal(String.IsNullOrEmpty(namedWorksheet.Cells[mn - 1, 37].Value.ToString()) ? "0" : namedWorksheet.Cells[mn - 1, 37].Value.ToString());
                                    namedWorksheet.Cells[mn, 38].Value = Convert.ToDecimal(String.IsNullOrEmpty(namedWorksheet.Cells[mn - 1, 38].Value.ToString()) ? "0" : namedWorksheet.Cells[mn - 1, 38].Value.ToString());
                                    var T_IRES = Convert.ToDecimal(String.IsNullOrEmpty(namedWorksheet.Cells[mn - 1, 37].Value.ToString()) ? "0" : namedWorksheet.Cells[mn - 1, 37].Value.ToString());

                                    var Str6 = "SELECT dbo.fn_FinalCountPN(" + "'" + T_CNTPN.ToString().Trim() + "'" + ")";

                                    using (SqlCommand cmd6 = new SqlCommand(Str6, sqlcon))
                                    {
                                        var result6 = cmd6.ExecuteScalar();
                                        //namedWorksheet.Cells[cntRow + 1, 31].Value = result;
                                        if (T_Value > 0)
                                        {
                                            if (Convert.ToInt32(result6.ToString()) == 1)
                                            {
                                                if (T_IRES > 0)
                                                {
                                                    namedWorksheet.Cells[mn, 33].Value = 'C';
                                                }
                                                else if (T_IRES <= 0)
                                                {
                                                    namedWorksheet.Cells[mn, 33].Value = 'W';
                                                }

                                            }
                                            if (Convert.ToInt32(result6.ToString()) == 0)
                                            {
                                                if (T_IRES > 0)
                                                {
                                                    namedWorksheet.Cells[mn, 33].Value = 'W';
                                                }
                                                else if (T_IRES <= 0)
                                                {
                                                    namedWorksheet.Cells[mn, 33].Value = 'C';
                                                }

                                            }
                                            if (Convert.ToInt32(result6.ToString()) == 2)
                                            {
                                                if (T_IRES > 0)
                                                {
                                                    namedWorksheet.Cells[mn, 33].Value = 'C';
                                                }
                                                else if (T_IRES <= 0)
                                                {
                                                    namedWorksheet.Cells[mn, 33].Value = 'W';
                                                }

                                            }
                                        }
                                        else if (T_Value < 0)
                                        {
                                            if (Convert.ToInt32(result6.ToString()) == 0)
                                            {
                                                if (T_IRES <= 0)
                                                {
                                                    namedWorksheet.Cells[mn, 33].Value = 'C';
                                                }
                                                else if (T_IRES > 0)
                                                {
                                                    namedWorksheet.Cells[mn, 33].Value = 'W';
                                                }

                                            }
                                            if (Convert.ToInt32(result6.ToString()) == 1)
                                            {
                                                if (T_IRES <= 0)
                                                {
                                                    namedWorksheet.Cells[mn, 33].Value = 'W';
                                                }
                                                else if (T_IRES > 0)
                                                {
                                                    namedWorksheet.Cells[mn, 33].Value = 'C';
                                                }

                                            }
                                            if (Convert.ToInt32(result6.ToString()) == 2)
                                            {
                                                if (T_IRES <= 0)
                                                {
                                                    namedWorksheet.Cells[mn, 33].Value = 'C';
                                                }
                                                else if (T_IRES > 0)
                                                {
                                                    namedWorksheet.Cells[mn, 33].Value = 'W';
                                                }

                                            }
                                        }


                                        //3col

                                        if (T_3COLValue > 0)
                                        {
                                            if (Convert.ToInt32(result6.ToString()) == 1)
                                            {
                                                if (T_IRES > 0)
                                                {
                                                    namedWorksheet.Cells[mn, 35].Value = 'C';
                                                }
                                                else if (T_IRES <= 0)
                                                {
                                                    namedWorksheet.Cells[mn, 35].Value = 'W';
                                                }

                                            }
                                            if (Convert.ToInt32(result6.ToString()) == 0)
                                            {
                                                if (T_IRES > 0)
                                                {
                                                    namedWorksheet.Cells[mn, 35].Value = 'W';
                                                }
                                                else if (T_IRES <= 0)
                                                {
                                                    namedWorksheet.Cells[mn, 35].Value = 'C';
                                                }

                                            }
                                            if (Convert.ToInt32(result6.ToString()) == 2)
                                            {
                                                if (T_IRES > 0)
                                                {
                                                    namedWorksheet.Cells[mn, 35].Value = 'C';
                                                }
                                                else if (T_IRES <= 0)
                                                {
                                                    namedWorksheet.Cells[mn, 35].Value = 'W';
                                                }

                                            }
                                        }
                                        else if (T_3COLValue < 0)
                                        {
                                            if (Convert.ToInt32(result6.ToString()) == 0)
                                            {
                                                if (T_IRES <= 0)
                                                {
                                                    namedWorksheet.Cells[mn, 35].Value = 'C';
                                                }
                                                else if (T_IRES > 0)
                                                {
                                                    namedWorksheet.Cells[mn, 35].Value = 'W';
                                                }

                                            }
                                            if (Convert.ToInt32(result6.ToString()) == 1)
                                            {
                                                if (T_IRES <= 0)
                                                {
                                                    namedWorksheet.Cells[mn, 35].Value = 'W';
                                                }
                                                else if (T_IRES > 0)
                                                {
                                                    namedWorksheet.Cells[mn, 35].Value = 'C';
                                                }

                                            }
                                            if (Convert.ToInt32(result6.ToString()) == 2)
                                            {
                                                if (T_IRES <= 0)
                                                {
                                                    namedWorksheet.Cells[mn, 35].Value = 'C';
                                                }
                                                else if (T_IRES > 0)
                                                {
                                                    namedWorksheet.Cells[mn, 35].Value = 'W';
                                                }

                                            }
                                        }


                                    }
                                }
                                if (Convert.ToString(namedWorksheet.Cells[mn, 15].Value) == "TOTAL_PRCL->PPCL" || Convert.ToString(namedWorksheet.Cells[mn, 15].Value) == "TOTAL_PRCL/330IND->PPOP/PPCL")
                                {
                                    var T_Value = Convert.ToDecimal(namedWorksheet.Cells[mn, 30].Value.ToString());
                                    var T_3COLValue = Convert.ToDecimal(namedWorksheet.Cells[mn, 34].Value.ToString());
                                    var T_CNTPN = namedWorksheet.Cells[mn, 31].Value.ToString();
                                    //MessageBox.Show(namedWorksheet.Cells[mn, 31].Value.ToString());
                                    //MessageBox.Show(namedWorksheet.Cells[mn - 1, 37].Value.ToString());
                                    namedWorksheet.Cells[mn, 37].Value = Convert.ToDecimal(String.IsNullOrEmpty(namedWorksheet.Cells[mn - 1, 37].Value.ToString()) ? "0" : namedWorksheet.Cells[mn - 1, 37].Value.ToString());
                                    namedWorksheet.Cells[mn, 38].Value = Convert.ToDecimal(String.IsNullOrEmpty(namedWorksheet.Cells[mn - 1, 38].Value.ToString()) ? "0" : namedWorksheet.Cells[mn - 1, 38].Value.ToString());

                                    var T_IRES = Convert.ToDecimal(String.IsNullOrEmpty(namedWorksheet.Cells[mn - 1, 37].Value.ToString()) ? "0" : namedWorksheet.Cells[mn - 1, 37].Value.ToString());

                                    var Str6 = "SELECT dbo.fn_FinalCountPN(" + "'" + T_CNTPN.ToString().Trim() + "'" + ")";

                                    using (SqlCommand cmd6 = new SqlCommand(Str6, sqlcon))
                                    {
                                        var result6 = cmd6.ExecuteScalar();
                                        //namedWorksheet.Cells[cntRow + 1, 31].Value = result;

                                        if (T_Value > 0)
                                        {
                                            if (Convert.ToInt32(result6.ToString()) == 1)
                                            {
                                                if (T_IRES > 0)
                                                {
                                                    namedWorksheet.Cells[mn, 33].Value = 'C';
                                                }
                                                else if (T_IRES <= 0)
                                                {
                                                    namedWorksheet.Cells[mn, 33].Value = 'W';
                                                }

                                            }
                                            if (Convert.ToInt32(result6.ToString()) == 0)
                                            {
                                                if (T_IRES > 0)
                                                {
                                                    namedWorksheet.Cells[mn, 33].Value = 'W';
                                                }
                                                else if (T_IRES <= 0)
                                                {
                                                    namedWorksheet.Cells[mn, 33].Value = 'C';
                                                }

                                            }
                                            if (Convert.ToInt32(result6.ToString()) == 2)
                                            {
                                                if (T_IRES > 0)
                                                {
                                                    namedWorksheet.Cells[mn, 33].Value = 'C';
                                                }
                                                else if (T_IRES <= 0)
                                                {
                                                    namedWorksheet.Cells[mn, 33].Value = 'W';
                                                }

                                            }

                                        }
                                        else if (T_Value < 0)
                                        {
                                            if (Convert.ToInt32(result6.ToString()) == 0)
                                            {
                                                if (T_IRES <= 0)
                                                {
                                                    namedWorksheet.Cells[mn, 33].Value = 'C';
                                                }
                                                else if (T_IRES > 0)
                                                {
                                                    namedWorksheet.Cells[mn, 33].Value = 'W';
                                                }

                                            }
                                            if (Convert.ToInt32(result6.ToString()) == 1)
                                            {
                                                if (T_IRES <= 0)
                                                {
                                                    namedWorksheet.Cells[mn, 33].Value = 'W';
                                                }
                                                else if (T_IRES > 0)
                                                {
                                                    namedWorksheet.Cells[mn, 33].Value = 'C';
                                                }

                                            }
                                            if (Convert.ToInt32(result6.ToString()) == 2)
                                            {
                                                if (T_IRES <= 0)
                                                {
                                                    namedWorksheet.Cells[mn, 33].Value = 'C';
                                                }
                                                else if (T_IRES > 0)
                                                {
                                                    namedWorksheet.Cells[mn, 33].Value = 'W';
                                                }

                                            }
                                        }


                                        //3col

                                        if (T_3COLValue > 0)
                                        {
                                            if (Convert.ToInt32(result6.ToString()) == 1)
                                            {
                                                if (T_IRES > 0)
                                                {
                                                    namedWorksheet.Cells[mn, 35].Value = 'C';
                                                }
                                                else if (T_IRES <= 0)
                                                {
                                                    namedWorksheet.Cells[mn, 35].Value = 'W';
                                                }

                                            }
                                            if (Convert.ToInt32(result6.ToString()) == 0)
                                            {
                                                if (T_IRES > 0)
                                                {
                                                    namedWorksheet.Cells[mn, 35].Value = 'W';
                                                }
                                                else if (T_IRES <= 0)
                                                {
                                                    namedWorksheet.Cells[mn, 35].Value = 'C';
                                                }

                                            }
                                            if (Convert.ToInt32(result6.ToString()) == 2)
                                            {
                                                if (T_IRES > 0)
                                                {
                                                    namedWorksheet.Cells[mn, 35].Value = 'C';
                                                }
                                                else if (T_IRES <= 0)
                                                {
                                                    namedWorksheet.Cells[mn, 35].Value = 'W';
                                                }

                                            }
                                        }
                                        else if (T_3COLValue < 0)
                                        {
                                            if (Convert.ToInt32(result6.ToString()) == 0)
                                            {
                                                if (T_IRES <= 0)
                                                {
                                                    namedWorksheet.Cells[mn, 35].Value = 'C';
                                                }
                                                else if (T_IRES > 0)
                                                {
                                                    namedWorksheet.Cells[mn, 35].Value = 'W';
                                                }

                                            }
                                            if (Convert.ToInt32(result6.ToString()) == 1)
                                            {
                                                if (T_IRES <= 0)
                                                {
                                                    namedWorksheet.Cells[mn, 35].Value = 'W';
                                                }
                                                else if (T_IRES > 0)
                                                {
                                                    namedWorksheet.Cells[mn, 35].Value = 'C';
                                                }

                                            }
                                            if (Convert.ToInt32(result6.ToString()) == 2)
                                            {
                                                if (T_IRES <= 0)
                                                {
                                                    namedWorksheet.Cells[mn, 35].Value = 'C';
                                                }
                                                else if (T_IRES > 0)
                                                {
                                                    namedWorksheet.Cells[mn, 35].Value = 'W';
                                                }

                                            }
                                            
                                        }
                                        
                                    }
                                    
                                }
                                
                            }
                            //Save your file
                            excelPackage.Save();
                        }

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                        //wb.Save();
                        //wb.Close(true);
                        //excelApp.Quit();
                        //System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                    }
                    finally
                    {
                        //wb.Save();
                        //wb.Close(true);
                        //excelApp.Quit();
                        //System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                        //sqlcon.Close();
                    }

            }
                    catch(Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                        sqlcon.Close();
                    }
                    finally
                    {
                        sqlcon.Close();
                    }
            }

        private void btnExportCPrcnt_Click(object sender, EventArgs e)
        {
            try
            {
                sqlcon.Open();
                String var = "SP_C_PERCENTAGEP2";
                //MessageBox.Show(var);
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = sqlcon;
                cmd.CommandText = var;
                cmd.CommandType = CommandType.StoredProcedure;
                //cmd.Parameters.Add("@NoOfCells", SqlDbType.VarChar).Value = txtEnternoofcells.Text;
                //cmd.Parameters.Add("@ITEM_NO", SqlDbType.VarChar).Value = txtItemNoP2.Text;
                cmd.CommandTimeout = 36000;
                cmd.ExecuteNonQuery();
                MessageBox.Show("Calculation Completed Successfully.");
                SqlDataAdapter sda = new SqlDataAdapter("SELECT [RED/INC],[FLAG],[TOT],[3/COL],MinNoOfCells FROM [dbo].[ExportSummaryP2] ORDER BY [RED/INC] asc", sqlcon);
                DataSet ds = new DataSet();
                sda.Fill(ds, "ProjectHelathEntryResult");
                dgv3.DataSource = ds;
                dgv3.DataMember = "ProjectHelathEntryResult";

                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                using (ExcelPackage excelPackage = new ExcelPackage())
                {

                    //Set some properties of the Excel document
                    excelPackage.Workbook.Properties.Author = "Sayan";
                    excelPackage.Workbook.Properties.Title = "P2-SUMMARY-TABLE";
                    excelPackage.Workbook.Properties.Subject = "P2-SUMMARY-TABLE";
                    excelPackage.Workbook.Properties.Created = DateTime.Now;

                    //Create the WorkSheet
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("P2-ExportSummary");

                    for (int i = 0; i <= dgv3.Columns.Count - 1; i++)
                    {
                        worksheet.Cells[1, i + 1].Value = dgv3.Columns[i].HeaderText;
                    }

                    /*And the information of your data*/
                    for (int i = 0; i <= dgv3.RowCount - 1; i++)
                    {
                        for (int j = 0; j <= dgv3.ColumnCount - 1; j++)
                        {
                            DataGridViewCell cell = dgv3[j, i];
                            worksheet.Cells[i + 2, j + 1].Value = cell.Value;

                        }
                    }

                    //Save your file
                    FileInfo fi = new FileInfo(@"C:\ProjectHealthApplication\Export\P2-SUMMARY-TABLE.xlsx");
                    excelPackage.SaveAs(fi);
                }
                try
                {
                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                    ExcelPackage ExcelPkg = new ExcelPackage();

                    FileInfo fi = new FileInfo(@"C:\ProjectHealthApplication\Export\P2-SUMMARY-TABLE.xlsx");
                    using (ExcelPackage excelPackage = new ExcelPackage(fi))
                    {
                        //Get a WorkSheet by name. If the worksheet doesn't exist, throw an exeption
                        ExcelWorksheet namedWorksheet = excelPackage.Workbook.Worksheets["P2-ExportSummary"];

                        //If you don't know if a worksheet exists, you could use LINQ,
                        //So it doesn't throw an exception, but return null in case it doesn't find it
                        ExcelWorksheet anotherWorksheet =
                            excelPackage.Workbook.Worksheets.FirstOrDefault(x => x.Name == "P2-ExportSummary");


                        //ExcelRange Rng1= namedWorksheet.Cells[2,1, dgv3.RowCount,2]
                        //int j = 0;
                        for (int k = 0; k < dgv3.ColumnCount; k++)
                        {

                            using (ExcelRange Rng = namedWorksheet.Cells[1, 1, 1, k + 1])
                            {
                                namedWorksheet.Row(1).Height = 40;
                                namedWorksheet.Column(k + 1).Width = 6;
                                namedWorksheet.Column(k + 1).Style.WrapText = true;
                                namedWorksheet.Cells[1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                                //namedWorksheet.Cells[1, 2].Style.
                                namedWorksheet.Cells[1, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[1, 2].Style.Fill.BackgroundColor.SetColor(Color.Yellow);

                                namedWorksheet.Cells[1, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[1, 3].Style.Fill.BackgroundColor.SetColor(Color.DeepSkyBlue);

                                //namedWorksheet.Cells[1, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                //namedWorksheet.Cells[1, 4].Style.Fill.BackgroundColor.SetColor(Color.Yellow);

                                //namedWorksheet.Cells[1, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                //namedWorksheet.Cells[1, 6].Style.Fill.BackgroundColor.SetColor(Color.DeepSkyBlue);

                                //namedWorksheet.Cells[1, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                //namedWorksheet.Cells[1, 6].Style.Fill.BackgroundColor.SetColor(Color.DeepSkyBlue);
                                
                                namedWorksheet.View.FreezePanes(2, 1);
                                //Rng.AutoFitColumns();

                                Rng.Style.Font.Size = 8;
                                Rng.Style.Font.Bold = true;
                                Rng.Style.Font.Color.SetColor(Color.Red);
                                Rng.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                Rng.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                Rng.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                Rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                Rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            }

                            using (ExcelRange Rng = namedWorksheet.Cells[2, (k + 1), dgv3.RowCount, (k + 1)])
                            {

                                    if (k == 2)
                                    {
                                        Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        Rng.Style.Fill.BackgroundColor.SetColor(Color.Gold);

                                    }
                                    else if (k == 1)
                                    {
                                        Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        Rng.Style.Fill.BackgroundColor.SetColor(Color.Green);
                                    }
                                    else if (k == 5)
                                    {
                                        Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        Rng.Style.Fill.BackgroundColor.SetColor(Color.DeepSkyBlue);
                                    }
                                    /*
                                    else if (k == 6)
                                    {
                                        Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        Rng.Style.Fill.BackgroundColor.SetColor(Color.LightCyan);
                                    }
                                    */
                                    else
                                    {
                                        Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        Rng.Style.Fill.BackgroundColor.SetColor(Color.White);
                                    }
                                    namedWorksheet.Column(k + 1).Width = 6;
                                    namedWorksheet.Column(k + 1).Style.WrapText = true;
                                    Rng.Style.Font.Size = 8;
                                    Rng.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                    Rng.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                    Rng.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                    Rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                    Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    Rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                    //Rng.AutoFitColumns();

                            }

                          }
                            using (ExcelRange Rng = namedWorksheet.Cells[1, 1, 1, dgv3.ColumnCount])
                            {

                                Rng.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                Rng.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                Rng.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                Rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                            }

                        //Save your file
                        excelPackage.Save();
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                sqlcon.Close();
            }
            finally
            {
                sqlcon.Close();
            }
        }

        private void btnExportCPrcntNotP2_Click(object sender, EventArgs e)
        {
            try
            {
                sqlcon.Open();
                String var = "SP_C_PERCENTAGE";
                //MessageBox.Show(var);
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = sqlcon;
                cmd.CommandText = var;
                cmd.CommandType = CommandType.StoredProcedure;
                //cmd.Parameters.Add("@NoOfCells", SqlDbType.VarChar).Value = txtEnternoofcells.Text;
                //cmd.Parameters.Add("@ITEM_NO", SqlDbType.VarChar).Value = txtItemNoP2.Text;
                cmd.CommandTimeout = 36000;
                cmd.ExecuteNonQuery();
                MessageBox.Show("Calculation Completed Successfully.");
                SqlDataAdapter sda = new SqlDataAdapter("select [RED/INC],[.03/ACT],[.05/ACT],[.1/ACT],[.03/AC3],[.5/AC3],[.1/AC3],[TOT] from [dbo].[ExportSummary] order by (case when [RED/INC]='FULL-TOT' THEN 20 when [RED/INC]='FULL-TOT-3COL' THEN 21 ELSE LEN([RED/INC]) END) asc", sqlcon);
                DataSet ds = new DataSet();
                sda.Fill(ds, "ProjectHelathEntryResult");
                dgv3.DataSource = ds;
                dgv3.DataMember = "ProjectHelathEntryResult";

                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                using (ExcelPackage excelPackage = new ExcelPackage())
                {

                    //Set some properties of the Excel document
                    excelPackage.Workbook.Properties.Author = "Sayan";
                    excelPackage.Workbook.Properties.Title = "P1-SUMMARY-TABLE";
                    excelPackage.Workbook.Properties.Subject = "P1-SUMMARY-TABLE";
                    excelPackage.Workbook.Properties.Created = DateTime.Now;

                    //Create the WorkSheet
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("ExportSummary");

                    for (int i = 0; i <= dgv3.Columns.Count - 1; i++)
                    {
                        worksheet.Cells[1, i + 1].Value = dgv3.Columns[i].HeaderText;
                    }

                    /*And the information of your data*/
                    for (int i = 0; i <= dgv3.RowCount - 1; i++)
                    {
                        for (int j = 0; j <= dgv3.ColumnCount - 1; j++)
                        {
                            DataGridViewCell cell = dgv3[j, i];
                            worksheet.Cells[i + 2, j + 1].Value = cell.Value;

                        }
                    }

                    //Save your file
                    FileInfo fi = new FileInfo(@"C:\ProjectHealthApplication\Export\P1-SUMMARY-TABLE.xlsx");
                    excelPackage.SaveAs(fi);
                }
                try
                {
                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                    ExcelPackage ExcelPkg = new ExcelPackage();

                    FileInfo fi = new FileInfo(@"C:\ProjectHealthApplication\Export\P1-SUMMARY-TABLE.xlsx");
                    using (ExcelPackage excelPackage = new ExcelPackage(fi))
                    {
                        //Get a WorkSheet by name. If the worksheet doesn't exist, throw an exeption
                        ExcelWorksheet namedWorksheet = excelPackage.Workbook.Worksheets["ExportSummary"];

                        //If you don't know if a worksheet exists, you could use LINQ,
                        //So it doesn't throw an exception, but return null in case it doesn't find it
                        ExcelWorksheet anotherWorksheet =
                            excelPackage.Workbook.Worksheets.FirstOrDefault(x => x.Name == "ExportSummary");


                        //ExcelRange Rng1= namedWorksheet.Cells[2,1, dgv3.RowCount,2]
                        //int j = 0;
                        for (int k = 0; k < dgv3.ColumnCount; k++)
                        {

                            using (ExcelRange Rng = namedWorksheet.Cells[1, 1, 1, k + 1])
                            {
                                namedWorksheet.Row(1).Height = 40;
                                namedWorksheet.Column(k + 1).Width = 6;
                                namedWorksheet.Column(k + 1).Style.WrapText = true;
                                namedWorksheet.Cells[1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                                //namedWorksheet.Cells[1, 2].Style.
                                namedWorksheet.Cells[1, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[1, 2].Style.Fill.BackgroundColor.SetColor(Color.Yellow);

                                namedWorksheet.Cells[1, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[1, 3].Style.Fill.BackgroundColor.SetColor(Color.Gold);

                                namedWorksheet.Cells[1, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                namedWorksheet.Cells[1, 4].Style.Fill.BackgroundColor.SetColor(Color.Yellow);

                                //namedWorksheet.Cells[1, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                //namedWorksheet.Cells[1, 6].Style.Fill.BackgroundColor.SetColor(Color.DeepSkyBlue);

                                //namedWorksheet.Cells[1, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                //namedWorksheet.Cells[1, 6].Style.Fill.BackgroundColor.SetColor(Color.DeepSkyBlue);

                                namedWorksheet.View.FreezePanes(2, 1);
                                //Rng.AutoFitColumns();

                                Rng.Style.Font.Size = 8;
                                Rng.Style.Font.Bold = true;
                                Rng.Style.Font.Color.SetColor(Color.Red);
                                Rng.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                Rng.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                Rng.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                Rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                Rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            }

                            using (ExcelRange Rng = namedWorksheet.Cells[2, (k + 1), dgv3.RowCount, (k + 1)])
                            {

                                if (k == 2)
                                {
                                    Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    Rng.Style.Fill.BackgroundColor.SetColor(Color.Gold);

                                }
                                else if (k == 3)
                                {
                                    Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    Rng.Style.Fill.BackgroundColor.SetColor(Color.Green);
                                }
                                else if (k == 5)
                                {
                                    Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    Rng.Style.Fill.BackgroundColor.SetColor(Color.DeepSkyBlue);
                                }
                                /*
                                else if (k == 6)
                                {
                                    Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    Rng.Style.Fill.BackgroundColor.SetColor(Color.LightCyan);
                                }
                                */
                                else
                                {
                                    Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    Rng.Style.Fill.BackgroundColor.SetColor(Color.White);
                                }
                                namedWorksheet.Column(k + 1).Width = 6;
                                namedWorksheet.Column(k + 1).Style.WrapText = true;
                                Rng.Style.Font.Size = 8;
                                Rng.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                Rng.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                Rng.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                Rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                Rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                //Rng.AutoFitColumns();

                            }

                        }
                        using (ExcelRange Rng = namedWorksheet.Cells[1, 1, 1, dgv3.ColumnCount])
                        {

                            Rng.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            Rng.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            Rng.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            Rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        }

                        //Save your file
                        excelPackage.Save();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                sqlcon.Close();
            }
            finally
            {
                sqlcon.Close();
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            try
            {
                sqlcon.Open();
                String var = "SP_UpdateProjectHealthEntry";
                //MessageBox.Show(var);
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = sqlcon;
                cmd.CommandText = var;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@NoOfCells", SqlDbType.Int).Value = Convert.ToInt32(txtNoofCells.Text.ToString());
                //cmd.Parameters.Add("@ITEM_NO", SqlDbType.VarChar).Value = txtItemNoP2.Text;
                cmd.CommandTimeout = 3600000;
                cmd.ExecuteNonQuery();
                MessageBox.Show("No Of Cells Limitation Updated Successfully.");
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
