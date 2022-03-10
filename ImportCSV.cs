using LumenWorks.Framework.IO.Csv;
using System;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Text;
using System.Windows.Forms;


namespace Label_Print
{
    public partial class ImportCSV : Form
    {
        public ImportCSV()
        {
            InitializeComponent();
        }

        public static string path = Application.StartupPath;
        public static String[,] array = new String[2000, 18];
        public static int ligne = 0;

        private void btnSave_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

            try
            {
                DataTable dtItem = (DataTable)(dgvImportCSV.DataSource);
                string EPN, CPN1, CPN2, CPN3;
                string InsertItemQry = "";
                int count = 0;
                foreach (DataRow dr in dtItem.Rows)

                {
                    Console.WriteLine("here2");
                    EPN = Convert.ToString(dr["EPN"]);
                    CPN1 = Convert.ToString(dr["CPN1"]);
                    CPN2 = Convert.ToString(dr["CPN2"]);
                    CPN3 = Convert.ToString(dr["CPN3"]);
                    if (EPN != "" && CPN1 != "" && CPN2 != "" && CPN3 != "")
                    {
                        Console.WriteLine("here3");
                        Console.WriteLine("epn/" + dr["EPN"] + "/");
                        Console.WriteLine("cpn1/" + dr["CPN1"] + "/");
                        //InsertItemQry += "Insert into REFERENCE(LookupCode,ItemName,DeptId,Cost,Price, Quantity, UOM, Weight, TaxID, IsDiscountItem,EntryDate)Values('" + Lookup + "','" + description + "','" + dept + "','" + dr["Cat"] + "','" + dr["Cost"] + "','" + UnitPrice + "'," + dr["Quantity"] + ",'" + dr["UOM"] + "','" + dr["Weight"] + "','" + dr["TaxID"] + "','" + dr["IsDiscountItem"] + "',GETDATE()); ";
                        //InsertItemQry += "Insert into REFERENCE(EPN,CPN1,CPN2,CPN3,ALERT1,ALERT2,ALERT3,RELEASE,CUSTOMER,LOT,LEVEL_,INDICE,ETOILE,PRFX,ID_FAMILY,OPERATOR,OLL,ID_ETIQUETTE)Values('" + Lookup + "','" + description + "','" + dept + "','" + dr["CateId"] + "','" + dr["Cost"] + "','" + UnitPrice + "'," + dr["Quantity"] + ",'" + dr["UOM"] + "','" + dr["Weight"] + "','" + dr["TaxID"] + "','" + dr["IsDiscountItem"] + "',GETDATE()); ";
                        InsertItemQry += "Insert into REFERENCE(EPN,CPN1,CPN2,CPN3,ALERT1,ALERT2,ALERT3,RELEASE,CUSTOMER,LOT,LEVEL_,INDICE,ETOILE,PRFX,ID_FAMILY,OPERATOR,OLL,ID_ETIQUETTE)Values('" + dr["EPN"] + "','" + dr["CPN1"] + "','" + dr["CPN2"] + "','" + dr["CPN3"] + "','" + dr["ALERT1"] + "','" + dr["ALERT2"] + "','" + dr["ALERT3"] + "','" + dr["RELEASE"] + "','" + dr["CUSTOMER"] + "','" + dr["LOT"] + "','" + dr["LEVEL"] + "','" + dr["INDICE"] + "','" + dr["ETOILE"] + "','" + dr["PRFX"] + "','" + dr["FAMILY"] + "','" + dr["OPERATOR"] + "','" + dr["OLL"] + "','1')";
                        count++;
                    }
                }
                Console.WriteLine("here4");
                if (InsertItemQry.Length > 5)
                {
                    bool isSuccess = DBAccess.ExecuteQuery(InsertItemQry);
                    if (isSuccess)
                    {
                        Console.WriteLine("here5");
                        MessageBox.Show("Item Imported Successfully, Total Imported Records : " + count + "", "GAUTAM POS", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        dgvImportCSV.DataSource = null;
                    }
                }
                Console.WriteLine("here6");
            }
            catch (Exception ex)
            {
                Console.WriteLine("here7");
                MessageBox.Show("Exception :  " + ex);
            }
            Form1 form1 = new Form1();
            
            string Query = "SELECT [EPN] FROM[ETQSARCHIVE].[dbo].[REFERENCE]";
            form1.dgvRef.DataSource = DBAccess.ExecuteQuery2(Query);
        }

        private void btnImport_Click(object sender, EventArgs e)
        {

            try
            {
                File.ReadAllText(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\LabelPrint\\Template\\" + guna2ComboBox1.Text);
            }
            catch
            {
                MessageBox.Show("Please select a Template label");
                return;
            }

            try
            {
                string Quer = "SELECT * FROM[ETQSARCHIVE].[dbo].[ETIQUETTE] WHERE [NOM_ETI] ='" + guna2ComboBox1.Text + "'";
                var Table = new DataTable();
               
                DBAccess.FillDataTable(Quer, Table);
                string nomEti = Table.Rows[0][0].ToString();

                DataTable dtItem = (DataTable)(dgvImportCSV.DataSource);
                string EPN, CPN1, CPN2, CPN3;
                string InsertItemQry = "";
                int count = 0;
                int rowNum = 0;

                foreach (DataRow dr in dtItem.Rows)
                {
                    rowNum++;
                    EPN = Convert.ToString(dr["EPN"]);
                    if (EPN == "" || EPN == null)
                    {
                        MessageBox.Show("You need to add EPN in row :" + rowNum + "", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

                foreach (DataRow dr in dtItem.Rows)
                {
                    Console.WriteLine("here2");
                    EPN = Convert.ToString(dr["EPN"]);
                    CPN1 = Convert.ToString(dr["CPN1"]);
                    CPN2 = Convert.ToString(dr["CPN2"]);
                    CPN3 = Convert.ToString(dr["CPN3"]);
                    if (EPN != "" )
                    //if (EPN != "" && CPN1 != "" && CPN2 != "" && CPN3 != "")
                    {
                        Console.WriteLine("here3");
                        Console.WriteLine("epn/" + dr["EPN"] + "/");
                        Console.WriteLine("cpn1/" + dr["CPN1"] + "/");
                        InsertItemQry += "Insert into REFERENCE(EPN,CPN1,CPN2,CPN3,ALERT1,ALERT2,ALERT3,RELEASE,CUSTOMER,LOT,LEVEL_,INDICE,ETOILE,PRFX,ID_FAMILY,OPERATOR,OLL,ID_ETIQUETTE)Values('" + dr["EPN"] + "','" + dr["CPN1"] + "','" + dr["CPN2"] + "','" + dr["CPN3"] + "','" + dr["ALERT1"] + "','" + dr["ALERT2"] + "','" + dr["ALERT3"] + "','" + dr["RELEASE"] + "','" + dr["CUSTOMER"] + "','" + dr["LOT"] + "','" + dr["LEVEL"] + "','" + dr["INDICE"] + "','" + dr["ETOILE"] + "','" + dr["PRFX"] + "','" + dr["FAMILY"] + "','" + dr["OPERATOR"] + "','" + dr["OLL"] + "','"+ nomEti + "')";
                        count++;
                    }
                    else
                    {
                        MessageBox.Show("EPN, CPN1, CPN2 and CPN3 shold not be empty : " + count + "", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                Console.WriteLine("here4");
                if (InsertItemQry.Length > 5)
                {
                    bool isSuccess = DBAccess.ExecuteQuery(InsertItemQry);
                    if (isSuccess)
                    {
                        Console.WriteLine("here5");
                        MessageBox.Show("Item Imported Successfully, Total Imported Records : " + count + "", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        dgvImportCSV.DataSource = null;
                        Form1 form1 = new Form1();
                        string Query = "SELECT [EPN] FROM[ETQSARCHIVE].[dbo].[REFERENCE]";
                        form1.dgvRef.DataSource = DBAccess.ExecuteQuery2(Query);
                        this.Dispose();
                    }
                }
                Console.WriteLine("here6");
            }
            catch (Exception ex)
            {
                Console.WriteLine("here7");
                MessageBox.Show("Exception :  " + ex);
            }

        }

        private void ImportCSV_Load(object sender, EventArgs e)
        {

            DataTable table = new DataTable();

            DBAccess.FillDataTable("SELECT [NOM_ETI] FROM [ETQSARCHIVE].[dbo].[ETIQUETTE]", table);

            guna2ComboBox1.DataSource = table;
            //cbTemplate.DisplayMember = "File Name";
            guna2ComboBox1.ValueMember = "NOM_ETI";

        }
        public static DataTable GetDataTabletFromCSVFile(string csv_file_path)
        {
            DataTable csvData = new DataTable();
            try
            {
                if (csv_file_path.EndsWith(".csv"))
                {
                    using (Microsoft.VisualBasic.FileIO.TextFieldParser csvReader = new Microsoft.VisualBasic.FileIO.TextFieldParser(csv_file_path))
                    {
                        csvReader.SetDelimiters(new string[] { ";" });
                        csvReader.HasFieldsEnclosedInQuotes = true;
                        //read column
                        string[] colFields = csvReader.ReadFields();
                        foreach (string column in colFields)
                        {
                            DataColumn datecolumn = new DataColumn(column);
                            datecolumn.AllowDBNull = true;
                            csvData.Columns.Add(datecolumn);
                        }
                        while (!csvReader.EndOfData)
                        {
                            string[] fieldData = csvReader.ReadFields();
                            for (int i = 0; i < fieldData.Length; i++)
                            {
                                if (fieldData[i] == "")
                                {
                                    fieldData[i] = null;
                                }
                            }
                            csvData.Rows.Add(fieldData);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exce " + ex);
            }
            return csvData;
        }

        private void btnChooseCSV_Click(object sender, EventArgs e)
        {
            string fileName = null;
            //System.Threading.Thread.Sleep(1500);
            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                openFileDialog1.InitialDirectory = path;
                openFileDialog1.Filter = "csv files (*.csv)|*.csv";
                openFileDialog1.FilterIndex = 2;
                openFileDialog1.RestoreDirectory = true;

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    fileName = openFileDialog1.FileName;
                    Console.WriteLine("path: " + fileName);
                }

                try
                {
                    int ImportedRecord = 0, inValidItem = 0;
                    string SourceURl = "";

                    if (fileName != null)
                    {
                        if (fileName.EndsWith(".csv"))
                        {
                            DataTable dtNew = new DataTable();
                            dtNew = GetDataTabletFromCSVFile(fileName);
                            if (Convert.ToString(dtNew.Columns[0]).ToUpper() != "EPN" 
                                || Convert.ToString(dtNew.Columns[1]).ToUpper() != "CPN1"
                                || Convert.ToString(dtNew.Columns[2]).ToUpper() != "CPN2"
                                || Convert.ToString(dtNew.Columns[3]).ToUpper() != "CPN3"
                                || Convert.ToString(dtNew.Columns[4]).ToUpper() != "ALERT1"
                                || Convert.ToString(dtNew.Columns[5]).ToUpper() != "ALERT2"
                                || Convert.ToString(dtNew.Columns[6]).ToUpper() != "ALERT3"
                                || Convert.ToString(dtNew.Columns[7]).ToUpper() != "RELEASE"
                                || Convert.ToString(dtNew.Columns[8]).ToUpper() != "CUSTOMER"
                                || Convert.ToString(dtNew.Columns[9]).ToUpper() != "FAMILY"
                                || Convert.ToString(dtNew.Columns[10]).ToUpper() != "OPERATOR"
                                || Convert.ToString(dtNew.Columns[11]).ToUpper() != "OLL"
                                || Convert.ToString(dtNew.Columns[12]).ToUpper() != "LOT"
                                || Convert.ToString(dtNew.Columns[13]).ToUpper() != "LEVEL"
                                || Convert.ToString(dtNew.Columns[14]).ToUpper() != "INDICE"
                                || Convert.ToString(dtNew.Columns[15]).ToUpper() != "ETOILE"
                                || Convert.ToString(dtNew.Columns[16]).ToUpper() != "PRFX"
                                || Convert.ToString(dtNew.Columns[17]).ToUpper() != "QUANTITE"
                                )
                            {
                                DialogResult dialogResult = MessageBox.Show("The template is invalid \n\nDo you want a template sample ?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                                if (dialogResult == DialogResult.Yes)
                                {
                                    System.Diagnostics.Process.Start(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\LabelPrint\\MISE EN FORME.csv");
                                }

                                return;
                            }

                            SourceURl = fileName;
                            if (dtNew.Rows != null && dtNew.Rows.ToString() != String.Empty)
                            {
                                dgvImportCSV.DataSource = dtNew;
                            }
                            foreach (DataGridViewRow row in dgvImportCSV.Rows)
                            {
                                if (Convert.ToString(row.Cells["EPN"].Value) == "" || row.Cells["EPN"].Value == null)
                                {
                                    row.DefaultCellStyle.BackColor = Color.Red;
                                    inValidItem += 1;
                                }
                                else
                                {
                                    ImportedRecord += 1;
                                }

                                btnPrintCSV.Enabled = true;
                                btnImport.Enabled = true;
                            }
                            Console.WriteLine("dgvImportCSV.Rows.Count " + dgvImportCSV.Rows.Count);
                            if (dgvImportCSV.Rows.Count == 1)
                            {
                                MessageBox.Show("There is no data in this file", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                        }
                        else
                        {
                            MessageBox.Show("Selected File is Invalid, Please Select valid csv file.", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Exception " + ex);
                    Console.WriteLine("here ----------------");
                }

            }
        }


        private void btnPrintCSV_Click(object sender, EventArgs e)
        {

            try
            {
                File.ReadAllText(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\LabelPrint\\Template\\" + guna2ComboBox1.Text);
            }
            catch
            {
                MessageBox.Show("Please select a Template label");
                return;
            }

            Form1 form1 = new Form1();
            int year = form1.dtpicker.Value.Year;
            int month = form1.dtpicker.Value.Month;
            int day = form1.dtpicker.Value.Day;
            int hour = form1.dtpicker.Value.Hour;
            int minute = form1.dtpicker.Value.Minute;


            try
            {
                DataTable dtItem = (DataTable)(dgvImportCSV.DataSource);
                Random rnd = new Random();
                string EPN;

                int i = 0;
                int rowNum = 0;
                String printer = "";

                foreach (DataRow dr in dtItem.Rows)
                {
                    rowNum++;
                    EPN = Convert.ToString(dr["EPN"]);
                    if (EPN == "" || EPN == null)
                    {
                        MessageBox.Show("You need to add EPN in row :" + rowNum + "", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

                    foreach (DataRow dr in dtItem.Rows)
                    {
                    EPN = Convert.ToString(dr["EPN"]);
                    
                    if (EPN != "" && EPN != null)
                    {
                        
                        int copy = Convert.ToInt32(dr["QUANTITE"]);

                        System.Threading.Thread.Sleep(1500);

                       
                        if (i == 0)
                        {
                            PrintDialog pd = new PrintDialog();

                            pd.PrinterSettings = new PrinterSettings();
                            if (DialogResult.OK == pd.ShowDialog(this))
                            {
                                for (int cp = 0; cp < copy; cp++)
                                {
                                    string str = File.ReadAllText(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\LabelPrint\\Template\\" + guna2ComboBox1.Text);
                                    str = str.Replace("@EPN", Convert.ToString(dr["EPN"]));
                                    str = str.Replace("@CPN1", Convert.ToString(dr["CPN1"]));
                                    str = str.Replace("@CPN2", Convert.ToString(dr["CPN2"]));
                                    str = str.Replace("@CPN3", Convert.ToString(dr["CPN3"]));
                                    str = str.Replace("@ALERT1", Convert.ToString(dr["ALERT1"]));
                                    str = str.Replace("@ALERT2", Convert.ToString(dr["ALERT2"]));
                                    str = str.Replace("@ALERT3", Convert.ToString(dr["ALERT3"]));
                                    str = str.Replace("@RELEASE", Convert.ToString(dr["RELEASE"]));
                                    str = str.Replace("@FIRST_CUSTOMER", Convert.ToString(dr["CUSTOMER"]));
                                    str = str.Replace("@FAMILLE", Convert.ToString(dr["FAMILY"]));
                                    str = str.Replace("@OPERID", Convert.ToString(dr["OPERATOR"]));
                                    str = str.Replace("@OLL", Convert.ToString(dr["OLL"]));
                                    str = str.Replace("@LOT", Convert.ToString(dr["LOT"]));
                                    str = str.Replace("@LEVEL", Convert.ToString(dr["LEVEL"]));
                                    str = str.Replace("@INDICE", Convert.ToString(dr["INDICE"]));
                                    str = str.Replace("@ETOILE", Convert.ToString(dr["ETOILE"]));
                                    str = str.Replace("@PRFX", Convert.ToString(dr["PRFX"]));

                                    str = str.Replace("@JJ", day.ToString());
                                    str = str.Replace("@MM", month.ToString());
                                    str = str.Replace("@YY", year.ToString());
                                    str = str.Replace("@HOURS", form1.dtpicker.Value.ToString("hh:mm"));
                                  
                                    str = str.Replace("@COUNTER", rnd.Next(1000, 9999).ToString());

                                    if (form1.jjmmyyyy.Checked == true)
                                    {
                                        str = str.Replace("@DATETIME", form1.dtpicker.Value.ToString("dd/MM/yyyy"));
                                        str = str.Replace("@YY", year.ToString());
                                    }
                                    else
                                    {
                                        str = str.Replace("@DATETIME", form1.dtpicker.Value.ToString("dd/MM/yy"));
                                        str = str.Replace("@YY", year.ToString().Substring(2));
                                    }


                                    RawPrinterHelper.SendStringToPrinter(pd.PrinterSettings.PrinterName, str);
                                    form1.dtpicker.Value = form1.dtpicker.Value.AddMinutes(2);

                                }
                                printer = pd.PrinterSettings.PrinterName;
                            }
                        }
                        else
                        {
                            for (int cp = 0; cp < copy; cp++)
                            {
                                string str = File.ReadAllText(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\LabelPrint\\Template\\" + guna2ComboBox1.Text);
                                str = str.Replace("@EPN", Convert.ToString(dr["EPN"]));
                                str = str.Replace("@CPN1", Convert.ToString(dr["CPN1"]));
                                str = str.Replace("@CPN2", Convert.ToString(dr["CPN2"]));
                                str = str.Replace("@CPN3", Convert.ToString(dr["CPN3"]));
                                str = str.Replace("@ALERT1", Convert.ToString(dr["ALERT1"]));
                                str = str.Replace("@ALERT2", Convert.ToString(dr["ALERT2"]));
                                str = str.Replace("@ALERT3", Convert.ToString(dr["ALERT3"]));
                                str = str.Replace("@RELEASE", Convert.ToString(dr["RELEASE"]));
                                str = str.Replace("@FIRST_CUSTOMER", Convert.ToString(dr["CUSTOMER"]));
                                str = str.Replace("@FAMILLE", Convert.ToString(dr["FAMILY"]));
                                str = str.Replace("@OPERID", Convert.ToString(dr["OPERATOR"]));
                                str = str.Replace("@OLL", Convert.ToString(dr["OLL"]));
                                str = str.Replace("@LOT", Convert.ToString(dr["LOT"]));
                                str = str.Replace("@LEVEL", Convert.ToString(dr["LEVEL"]));
                                str = str.Replace("@INDICE", Convert.ToString(dr["INDICE"]));
                                str = str.Replace("@ETOILE", Convert.ToString(dr["ETOILE"]));
                                str = str.Replace("@PRFX", Convert.ToString(dr["PRFX"]));

                                str = str.Replace("@JJ", day.ToString());
                                str = str.Replace("@MM", month.ToString());
                                str = str.Replace("@YY", year.ToString());
                                str = str.Replace("@HOURS", form1.dtpicker.Value.ToString("hh:mm"));
                                
                                
                                str = str.Replace("@COUNTER", rnd.Next(1000, 8000).ToString());

                                if (form1.jjmmyyyy.Checked == true)
                                {
                                    str = str.Replace("@DATETIME", form1.dtpicker.Value.ToString("dd/MM/yyyy"));
                                    str = str.Replace("@YY", year.ToString());
                                }
                                else
                                {
                                    str = str.Replace("@DATETIME", form1.dtpicker.Value.ToString("dd/MM/yy"));
                                    str = str.Replace("@YY", year.ToString().Substring(2));
                                }
                               
                                RawPrinterHelper.SendStringToPrinter(printer, str);
                                form1.dtpicker.Value = form1.dtpicker.Value.AddMinutes(2);
                            }
                        }

                        File.Delete(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\LabelPrint\\tmp.prn");
                        i++;
                    }
                    else
                    {
                        MessageBox.Show("You need to add EPN in row :" + dtItem.Rows.Count + "", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("here7");
                MessageBox.Show("Exception :  " + ex.Message);
            }

        }
    }
}
