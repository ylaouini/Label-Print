using LumenWorks.Framework.IO.Csv;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;


namespace Label_Print
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public static string path = Application.StartupPath;
        public static String[,] array = new String[2000, 18];
        public static int ligne = 0;
        public string t;

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

            string path = Application.StartupPath + "/Template";
            DataTable table = new DataTable();
            table.Columns.Add("File Name");
            table.Columns.Add("File Path");

            string[] files = Directory.GetFiles(path);

            for (int i = 0; i < files.Length; i++)
            {
                FileInfo file = new FileInfo(files[i]);
                table.Rows.Add(file.Name, path + "\\" + file.Name);
            }


            cbTemplate.DataSource = table;
            cbTemplate.DisplayMember = "File Name";
            cbTemplate.ValueMember = "File Path";
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            Console.WriteLine(dtpicker.Value.AddMinutes(2));
            
            int year = dtpicker.Value.Year;
            int month = dtpicker.Value.Month;
            int day = dtpicker.Value.Day;
            int hour = dtpicker.Value.Hour;
            int minute = dtpicker.Value.Minute;
            
            Console.WriteLine("Year :" + year +" Day : "+ day + " month :" +month);
            Console.WriteLine("hour :" + hour + " Minute : " + minute );
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dtpicker.Height = 40;

            DataTable table = new DataTable();

            DBAccess.FillDataTable("SELECT [NOM_ETI] FROM [ETQSARCHIVE].[dbo].[ETIQUETTE]", table);

            //Console.WriteLine("test " + table.Rows[0][0].ToString());

            cbTemplate.DataSource = table;
            //cbTemplate.DisplayMember = "File Name";
            cbTemplate.ValueMember = "NOM_ETI";

            //string Query = "SELECT [EPN] FROM[ETQSARCHIVE].[dbo].[REFERENCE]";
            //guna2DataGridView1.DataSource = DBAccess.ExecuteQuery2("SELECT [EPN] FROM[ETQSARCHIVE].[dbo].[REFERENCE]");

            var csvTable = new DataTable();
         
            try
            {
                string Query = "SELECT [EPN] FROM [ETQSARCHIVE].[dbo].[REFERENCE]";
                DBAccess.FillDataTable(Query,csvTable);
                dgvRef.DataSource = csvTable;
                DataGridViewColumn column = dgvRef.Columns[0];
                column.Width = 190;
            }
            catch(Exception ex)
            {
               MessageBox.Show(ex.Message);
            }
        }

        private void dgvRef_Click(object sender, EventArgs e)
        {
            string Query = "SELECT * FROM[ETQSARCHIVE].[dbo].[REFERENCE] WHERE [EPN] ='" + dgvRef.SelectedCells[0].Value + "'";
            var Table = new DataTable();
           
            try
            {
                btnPrint.Enabled = true;
                btnEdition.Enabled = true;
                DBAccess.FillDataTable(Query, Table);

                txtEPN.Text = Table.Rows[0][1].ToString();
                txtCPN1.Text = Table.Rows[0][2].ToString();
                txtCPN2.Text = Table.Rows[0][3].ToString();
                txtCPN3.Text = Table.Rows[0][4].ToString();
                txtALERT1.Text = Table.Rows[0][5].ToString();
                txtALERT2.Text = Table.Rows[0][6].ToString();
                txtALERT3.Text = Table.Rows[0][7].ToString();
                txtRELEASE.Text = Table.Rows[0][8].ToString();
                txtFCOSTUMER.Text = Table.Rows[0][9].ToString();

                txtLOT.Text = Table.Rows[0][10].ToString();
                txtLEVEL.Text = Table.Rows[0][11].ToString();
                txtINDICE.Text = Table.Rows[0][12].ToString();
                txtETOILE.Text = Table.Rows[0][13].ToString();
                txtPRFX.Text = Table.Rows[0][14].ToString();
                txtFAMILY.Text = Table.Rows[0][15].ToString();
                txtMATRICULE.Text = Table.Rows[0][16].ToString();
                txtOLL.Text = Table.Rows[0][17].ToString();
                 
                txtQUANTITE.Text = Table.Rows[0][18].ToString();
                 
            }
            catch (System.IO.IOException ex)
            {
                MessageBox.Show(ex.Message);

            }

            /**
            var csvTable = new DataTable();
            using (var csvReader = new CsvReader(new StreamReader(System.IO.File.OpenRead(Application.StartupPath + "/BD/bd.csv")), true))
            {
                csvTable.Load(csvReader);
            }

            for (int i = 0; i < csvTable.Rows.Count; i++)
            {

                if (csvTable.Rows[i][0].ToString().Equals(dgvRef.SelectedCells[0].Value) )
                {

                    txtEPN.Text = csvTable.Rows[i][0].ToString();
                    txtCPN1.Text = csvTable.Rows[i][1].ToString();
                    txtCPN2.Text = csvTable.Rows[i][2].ToString();
                    txtCPN3.Text = csvTable.Rows[i][3].ToString();
                    txtALERT1.Text = csvTable.Rows[i][4].ToString();
                    txtALERT2.Text = csvTable.Rows[i][5].ToString();
                    txtALERT3.Text = csvTable.Rows[i][6].ToString();
                    txtRELEASE.Text = csvTable.Rows[i][7].ToString();
                    txtFCOSTUMER.Text = csvTable.Rows[i][8].ToString();
                    txtFAMILY.Text = csvTable.Rows[i][9].ToString();
                    txtMATRICULE.Text = csvTable.Rows[i][10].ToString();
                    txtOLL.Text = csvTable.Rows[i][11].ToString();
                    txtPRFX.Text = csvTable.Rows[i][12].ToString();
                    txtLOT.Text = csvTable.Rows[i][13].ToString();
                    txtLEVEL.Text = csvTable.Rows[i][14].ToString();
                    txtQUANTITE.Text = csvTable.Rows[i][15].ToString();
                    txtINDICE.Text = csvTable.Rows[i][16].ToString();
                    txtETOILE.Text = csvTable.Rows[i][17].ToString();
                    return;
                }
            }*/

        }


        private void txtSearch_TextChanged(object sender, EventArgs e)
        {

            if(txtSearch.Text != "" && txtSearch.Text != "Search ...")
            {
                Console.WriteLine("here1");
                    string Query = "SELECT [EPN] FROM[ETQSARCHIVE].[dbo].[REFERENCE] WHERE EPN LIKE '%" + txtSearch.Text +"%'";

                    var csvTable = new DataTable();
                    try
                    {
                        DBAccess.FillDataTable(Query, csvTable);
                        dgvRef.DataSource = csvTable;
                    }
                    catch (System.IO.IOException ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
            }
            else if( txtSearch.Text != "Search ...")
            {
                Console.WriteLine("here2");
                var csvTable = new DataTable();
                string Query = "SELECT [EPN] FROM[ETQSARCHIVE].[dbo].[REFERENCE]";
                try
                {
                    DBAccess.FillDataTable(Query, csvTable);
                    dgvRef.DataSource = csvTable;
                }
                catch (System.IO.IOException ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            

            /*
            var csvTable = new DataTable();
            dgvRef.Rows.Clear();
            using (var csvReader = new CsvReader(new StreamReader(System.IO.File.OpenRead(Application.StartupPath + "/BD/bd.csv")), true))
            {
                csvTable.Load(csvReader);
            }


            for (int i = 0; i < csvTable.Rows.Count; i++)
            {
                if (csvTable.Rows[i][0].ToString().StartsWith(txtSearch.Text.ToUpper()) )
                {
                    dgvRef.Rows.Add(csvTable.Rows[i][0].ToString());
                }
            }*/

        }

        private void button1_Click_2(object sender, EventArgs e)
        {
            
            if(txtEPN.Text == "" || txtEPN.Text == null || txtEPN.Text == "EPN")
            {
                return;
            }

            if (btnEdition.Text.Equals("UPDATE")){

                if (txtEPN.Text == "EPN") { txtEPN.Text = ""; }
                if (txtCPN1.Text == "CPN 1") { txtCPN1.Text = ""; }
                if (txtCPN2.Text == "CPN 2") { txtCPN2.Text = ""; }
                if (txtCPN3.Text == "CPN 3") { txtCPN3.Text = ""; }
                if (txtALERT1.Text == "ALERT 1") { txtALERT1.Text = ""; }
                if (txtALERT2.Text == "ALERT 2") { txtALERT2.Text = ""; }
                if (txtALERT3.Text == "ALERT 3") { txtALERT3.Text = ""; }
                if (txtRELEASE.Text == "RELEASE") { txtRELEASE.Text = ""; }
                if (txtFCOSTUMER.Text == "FIRST CUSTOMER") { txtFCOSTUMER.Text = ""; }
                if (txtFAMILY.Text == "FAMILY") { txtFAMILY.Text = ""; }
                if (txtMATRICULE.Text == "OPERATOR ID") { txtMATRICULE.Text = ""; }
                if (txtLOT.Text == "LOT") { txtLOT.Text = ""; }
                if (txtLEVEL.Text == "LEVEL") { txtLEVEL.Text = ""; }
                if (txtINDICE.Text == "INDICE") { txtINDICE.Text = ""; }
                if (txtETOILE.Text == "ETOILE") { txtETOILE.Text = ""; }
                if (txtPRFX.Text == "PREFIX") { txtPRFX.Text = ""; }
                if (txtOLL.Text == "OLL") { txtOLL.Text = ""; }

                string Query = "UPDATE [dbo].[REFERENCE] SET [CPN1] = '" + txtCPN1.Text + "',[CPN2] = '" + txtCPN2.Text + "',[CPN3] = '" + txtCPN3.Text + "' ,[ALERT1] = '" + txtALERT1.Text + "',[ALERT2] = '" + txtALERT2.Text + "',[ALERT3] = '" + txtALERT3.Text + "',[RELEASE] ='" + txtRELEASE.Text + "',[CUSTOMER] = '" + txtFCOSTUMER.Text + "',[LOT] = '" + txtLOT.Text + "',[LEVEL_] = '" + txtLEVEL.Text + "' ,[INDICE] = '" + txtINDICE.Text + "' ,[ETOILE] = '" + txtETOILE.Text + "' ,[PRFX] = '" + txtPRFX.Text + "' ,[ID_FAMILY] = '" + txtFAMILY.Text + "' ,[OPERATOR] = '" + txtMATRICULE.Text + "' ,[OLL] = '" + txtOLL.Text + "' ,[ID_ETIQUETTE] = '' WHERE [EPN] = '" + txtEPN.Text + "'";

                bool isSuccess = DBAccess.ExecuteQuery(Query);
                Console.WriteLine(isSuccess + " / " + Query);
                if (isSuccess)
                {
                    
                    MessageBox.Show("Item Imported Successfully, Total Imported Records : ", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    

                    txtCPN1.Enabled = false;
                    txtCPN2.Enabled = false;
                    txtCPN3.Enabled = false;
                    txtALERT1.Enabled = false;
                    txtALERT2.Enabled = false;
                    txtALERT3.Enabled = false;
                    txtRELEASE.Enabled = false;
                    txtFCOSTUMER.Enabled = false;
                    txtFAMILY.Enabled = false;
                    txtMATRICULE.Enabled = false;
                    txtOLL.Enabled = false;
                    txtPRFX.Enabled = false;
                    txtLOT.Enabled = false;
                    txtLEVEL.Enabled = false;
                    txtQUANTITE.Enabled = false;
                    txtINDICE.Enabled = false;
                    txtETOILE.Enabled = false;


                    txtEPN.Text = "EPN";
                    txtCPN1.Text = "CPN 1";
                    txtCPN2.Text = "CPN 2";
                    txtCPN3.Text = "CPN 3";
                    txtALERT1.Text = "ALERT 1";
                    txtALERT2.Text = "ALERT 2";
                    txtALERT3.Text = "ALERT 3";
                    txtRELEASE.Text = "RELEASE";
                    txtFCOSTUMER.Text = "FIRST CUSTOMER";
                    txtFAMILY.Text = "FAMILY";
                    txtMATRICULE.Text = "OPERATOR ID";
                    txtLOT.Text = "LOT";
                    txtLEVEL.Text = "LEVEL";
                    txtINDICE.Text = "INDICE";
                    txtETOILE.Text = "ETOILE";
                    txtPRFX.Text = "PREFIX";
                    txtOLL.Text = "OLL";

                    btnEdition.Text = "EDITE";
                    btnNew.Enabled = true;
                    btnPrint.Enabled = false;
                    btnEdition.Enabled = false;
                    btnNew.Enabled = true;
                    dgvRef.Enabled = true;

                }


            }
            else
            {
                btnNew.Enabled = false;
                btnPrint.Enabled = false;
                dgvRef.Enabled = false;

                Console.WriteLine("edite "+btnEdition.Text);
                // txtEPN.Enabled = true;
                txtCPN1.Enabled = true;
                txtCPN2.Enabled = true;
                txtCPN3.Enabled = true;
                txtALERT1.Enabled = true;
                txtALERT2.Enabled = true;
                txtALERT3.Enabled = true;
                txtRELEASE.Enabled = true;
                txtFCOSTUMER.Enabled = true;
                txtFAMILY.Enabled = true;
                txtMATRICULE.Enabled = true;
                txtOLL.Enabled = true;
                txtPRFX.Enabled = true;
                txtLOT.Enabled = true;
                txtLEVEL.Enabled = true;
                txtQUANTITE.Enabled = true;
                txtINDICE.Enabled = true;
                txtETOILE.Enabled = true;

                txtCPN1.ForeColor = Color.Black;
                txtCPN2.ForeColor = Color.Black;
                txtCPN3.ForeColor = Color.Black;
                txtALERT1.ForeColor = Color.Black;
                txtALERT2.ForeColor = Color.Black;
                txtALERT3.ForeColor = Color.Black;
                txtRELEASE.ForeColor = Color.Black;
                txtFCOSTUMER.ForeColor = Color.Black;
                txtFAMILY.ForeColor = Color.Black;
                txtMATRICULE.ForeColor = Color.Black;
                txtOLL.ForeColor = Color.Black;
                txtPRFX.ForeColor = Color.Black;
                txtLOT.ForeColor = Color.Black;
                txtLEVEL.ForeColor = Color.Black;
                txtQUANTITE.ForeColor = Color.Black;
                txtINDICE.ForeColor = Color.Black;
                txtETOILE.ForeColor = Color.Black;
                txtALERT3.ForeColor = Color.Black;
                btnEdition.Text = "UPDATE";

            }
            txtEPN.Focus();
            txtCPN2.Focus();
            txtCPN3.Focus();
            txtALERT1.Focus();
            txtALERT2.Focus();
            txtALERT3.Focus();
            txtRELEASE.Focus();
            txtFCOSTUMER.Focus();
            txtFAMILY.Focus();
            txtMATRICULE.Focus();
            txtOLL.Focus();
            txtPRFX.Focus();
            txtLOT.Focus();
            txtLEVEL.Focus();
            txtQUANTITE.Focus();
            txtINDICE.Focus();
            txtETOILE.Focus();
            txtCPN1.Focus();
        }

       

        private static string FormatCSV(string input)
        {
            try
            {
                input = input.Replace("\"", "\"\"");
                return "\"" + input + "\"";
            }
            catch
            {
                throw;
            }
        }

        private void WriteCSVHeader(DataRow drMeta, string outputFilePath)
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                foreach (DataColumn dcMeta in drMeta.Table.Columns)
                {
                    sb.Append(FormatCSV(dcMeta.ColumnName.ToString()) + ",");
                }
                sb.Remove(sb.Length - 1, 1);
                sb.AppendLine();
                File.AppendAllText(outputFilePath, sb.ToString());
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        public void WriteCSV(DataRow drMeta, string outputFilePath)
        {
            try
            {
                StringBuilder sb = new StringBuilder();

                int cnt = File.ReadAllLines(outputFilePath).Length;

                #region write csv headers
                if (cnt == 0)
                {
                    WriteCSVHeader(drMeta, outputFilePath);
                }
                #endregion

                #region write csv rows
                foreach (DataColumn dcMeta in drMeta.Table.Columns)
                {
                    sb.Append(FormatCSV(drMeta[dcMeta.ColumnName].ToString()) + ",");
                }
                #endregion

                sb.Remove(sb.Length - 1, 1);
                sb.AppendLine();
                File.AppendAllText(outputFilePath, sb.ToString());
            }
            catch (Exception ex)
            {
                throw;
            }
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
                        csvReader.SetDelimiters(new string[] { "," });
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

        private void txtEPN_Enter(object sender, EventArgs e)
        {
            if(txtEPN.Text == "EPN")
            {
                txtEPN.Text = "";
                txtEPN.ForeColor = Color.Black;
            }
        }

        private void txtEPN_Leave(object sender, EventArgs e)
        {
            if (txtEPN.Text == "")
            {
                txtEPN.Text = "EPN";
                txtEPN.ForeColor = Color.Silver;
            }
        }

        private void txtCPN1_Enter(object sender, EventArgs e)
        {

            if (txtCPN1.Text == "CPN 1")
            {
                txtCPN1.Text = "";
                txtCPN1.ForeColor = Color.Black;
            }
        }

        private void txtCPN1_Leave(object sender, EventArgs e)
        {
            if (txtCPN1.Text == "")
            {
                txtCPN1.Text = "CPN 1";
                txtCPN1.ForeColor = Color.Silver;
            }
        }

        private void txtCPN2_Enter(object sender, EventArgs e)
        {
            if (txtCPN2.Text == "CPN 2")
            {
                txtCPN2.Text = "";
                txtCPN2.ForeColor = Color.Black;
            }
        }

        private void txtCPN2_Leave(object sender, EventArgs e)
        {
            if (txtCPN2.Text == "")
            {
                txtCPN2.Text = "CPN 2";
                txtCPN2.ForeColor = Color.Silver;
            }
        }

        private void txtCPN3_Enter(object sender, EventArgs e)
        {
            if (txtCPN3.Text == "CPN 3")
            {
                txtCPN3.Text = "";
                txtCPN3.ForeColor = Color.Black;
            }
        }

        private void txtCPN3_Leave(object sender, EventArgs e)
        {
            if (txtCPN3.Text == "")
            {
                txtCPN3.Text = "CPN 3";
                txtCPN3.ForeColor = Color.Silver;
            }
        }

        private void txtALERTE1_Enter(object sender, EventArgs e)
        {
            if (txtALERT1.Text == "ALERT 1")
            {
                txtALERT1.Text = "";
                txtALERT1.ForeColor = Color.Black;
            }
        }

        private void txtALERTE1_Leave(object sender, EventArgs e)
        {
            if (txtALERT1.Text == "")
            {
                txtALERT1.Text = "ALERT 1";
                txtALERT1.ForeColor = Color.Silver;
            }
        }

        private void txtRELEASE_Enter(object sender, EventArgs e)
        {
            if (txtRELEASE.Text == "RELEASE")
            {
                txtRELEASE.Text = "";
                txtRELEASE.ForeColor = Color.Black;
            }
        }

        private void txtRELEASE_Leave(object sender, EventArgs e)
        {
            if (txtRELEASE.Text == "")
            {
                txtRELEASE.Text = "RELEASE";
                txtRELEASE.ForeColor = Color.Silver;
            }
        }

        private void txtFCOSTUMER_Enter(object sender, EventArgs e)
        {
            if (txtFCOSTUMER.Text == "FIRST CUSTOMER")
            {
                txtFCOSTUMER.Text = "";
                txtFCOSTUMER.ForeColor = Color.Black;
            }
        }

        private void txtFCOSTUMER_Leave(object sender, EventArgs e)
        {
            if (txtFCOSTUMER.Text == "")
            {
                txtFCOSTUMER.Text = "FIRST CUSTOMER";
                txtFCOSTUMER.ForeColor = Color.Silver;
            }
        }

        private void txtFAMILY_Enter(object sender, EventArgs e)
        {
            if (txtFAMILY.Text == "FAMILY")
            {
                txtFAMILY.Text = "";
                txtFAMILY.ForeColor = Color.Black;
            }
        }

        private void txtFAMILY_Leave(object sender, EventArgs e)
        {
            if (txtFAMILY.Text == "")
            {
                txtFAMILY.Text = "FAMILY";
                txtFAMILY.ForeColor = Color.Silver;
            }
        }

        private void txtMATRICULE_Enter(object sender, EventArgs e)
        {
            if (txtMATRICULE.Text == "OPERATOR ID")
            {
                txtMATRICULE.Text = "";
                txtMATRICULE.ForeColor = Color.Black;
            }
        }

        private void txtMATRICULE_Leave(object sender, EventArgs e)
        {
            if (txtMATRICULE.Text == "")
            {
                txtMATRICULE.Text = "OPERATOR ID";
                txtMATRICULE.ForeColor = Color.Silver;
            }
        }

        private void txtOLL_Enter(object sender, EventArgs e)
        {
            if (txtOLL.Text == "OLL")
            {
                txtOLL.Text = "";
                txtOLL.ForeColor = Color.Black;
            }
        }

        private void txtOLL_Leave(object sender, EventArgs e)
        {
            if (txtOLL.Text == "")
            {
                txtOLL.Text = "OLL";
                txtOLL.ForeColor = Color.Silver;
            }
        }

        private void txtLOT_Enter(object sender, EventArgs e)
        {
            if (txtLOT.Text == "LOT")
            {
                txtLOT.Text = "";
                txtLOT.ForeColor = Color.Black;
            }
        }

        private void txtLOT_Leave(object sender, EventArgs e)
        {
            if (txtLOT.Text == "")
            {
                txtLOT.Text = "LOT";
                txtLOT.ForeColor = Color.Silver;
            }
        }

        private void txtLEVEL_Enter(object sender, EventArgs e)
        {
            if (txtLEVEL.Text == "LEVEL")
            {
                txtLEVEL.Text = "";
                txtLEVEL.ForeColor = Color.Black;
            }
        }

        private void txtLEVEL_Leave(object sender, EventArgs e)
        {
            if (txtLEVEL.Text == "")
            {
                txtLEVEL.Text = "LEVEL";
                txtLEVEL.ForeColor = Color.Silver;
            }
        }

        private void txtINDICE_Enter(object sender, EventArgs e)
        {
            if (txtINDICE.Text == "INDICE")
            {
                txtINDICE.Text = "";
                txtINDICE.ForeColor = Color.Black;
            }
        }

        private void txtINDICE_Leave(object sender, EventArgs e)
        {
            if (txtINDICE.Text == "")
            {
                txtINDICE.Text = "INDICE";
                txtINDICE.ForeColor = Color.Silver;
            }
        }

        private void txtETOILE_Enter(object sender, EventArgs e)
        {
            if (txtETOILE.Text == "ETOILE")
            {
                txtETOILE.Text = "";
                txtETOILE.ForeColor = Color.Black;
            }
        }

        private void txtETOILE_Leave(object sender, EventArgs e)
        {
            if (txtETOILE.Text == "")
            {
                txtETOILE.Text = "ETOILE";
                txtETOILE.ForeColor = Color.Silver;
            }
        }

        private void txtPRFX_Enter(object sender, EventArgs e)
        {
            if (txtPRFX.Text == "PREFIX")
            {
                txtPRFX.Text = "";
                txtPRFX.ForeColor = Color.Black;
            }
        }

        private void txtPRFX_Leave(object sender, EventArgs e)
        {
            if (txtPRFX.Text == "")
            {
                txtPRFX.Text = "PREFIX";
                txtPRFX.ForeColor = Color.Silver;
            }
        }

        private void txtALERT2_Enter(object sender, EventArgs e)
        {
            if (txtALERT2.Text == "ALERT 2")
            {
                txtALERT2.Text = "";
                txtALERT2.ForeColor = Color.Black;
            }
        }

        private void txtALERT2_Leave(object sender, EventArgs e)
        {
            if (txtALERT2.Text == "")
            {
                txtALERT2.Text = "ALERT 2";
                txtALERT2.ForeColor = Color.Silver;
            }
        }

        private void txtALERT3_Enter(object sender, EventArgs e)
        {
            if (txtALERT3.Text == "ALERT 3")
            {
                txtALERT3.Text = "";
                txtALERT3.ForeColor = Color.Black;
            }
        }

        private void txtALERT3_Leave(object sender, EventArgs e)
        {
            if (txtALERT3.Text == "")
            {
                txtALERT3.Text = "ALERT 3";
                txtALERT3.ForeColor = Color.Silver;
            }
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            if (btnNew.Text.Equals("INSERT"))
            {
                if (txtEPN.Text == "" || txtEPN.Text == null || txtEPN.Text == "EPN")
                {
                    MessageBox.Show("EPN cannot be empty ", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (txtEPN.Text == "EPN") { txtEPN.Text = ""; }
                if (txtCPN1.Text == "CPN 1") { txtCPN1.Text = ""; }
                if (txtCPN2.Text == "CPN 2") { txtCPN2.Text = ""; }
                if (txtCPN3.Text == "CPN 3") { txtCPN3.Text = ""; }
                if (txtALERT1.Text == "ALERT 1") { txtALERT1.Text = ""; }
                if (txtALERT2.Text == "ALERT 2") { txtALERT2.Text = ""; }
                if (txtALERT3.Text == "ALERT 3") { txtALERT3.Text = ""; }
                if (txtRELEASE.Text == "RELEASE") { txtRELEASE.Text = ""; }
                if (txtFCOSTUMER.Text == "FIRST CUSTOMER") { txtFCOSTUMER.Text = ""; }
                if (txtFAMILY.Text == "FAMILY") { txtFAMILY.Text = ""; }
                if (txtMATRICULE.Text == "OPERATOR ID") { txtMATRICULE.Text = ""; }
                if (txtLOT.Text == "LOT") { txtLOT.Text = ""; }
                if (txtLEVEL.Text == "LEVEL") { txtLEVEL.Text = ""; }
                if (txtINDICE.Text == "INDICE") { txtINDICE.Text = ""; }
                if (txtETOILE.Text == "ETOILE") { txtETOILE.Text = ""; }
                if (txtPRFX.Text == "PREFIX") { txtPRFX.Text = ""; }
                if (txtOLL.Text == "OLL") { txtOLL.Text = ""; }

                string Query = "INSERT INTO [dbo].[REFERENCE]  ([EPN] ,[CPN1] ,[CPN2] ,[CPN3]  ,[ALERT1] ,[ALERT2] ,[ALERT3] ,[RELEASE] ,[CUSTOMER] ,[LOT] ,[LEVEL_] ,[INDICE] ,[ETOILE] ,[PRFX] ,[ID_FAMILY] ,[OPERATOR] ,[OLL]) VALUES ('" + txtEPN.Text + "' , '" + txtCPN1.Text + "','" + txtCPN2.Text + "', '" + txtCPN3.Text + "','" + txtALERT1.Text + "', '" + txtALERT2.Text + "','" + txtALERT3.Text + "','" + txtRELEASE.Text + "','" + txtFCOSTUMER.Text + "','" + txtLOT.Text + "','" + txtLEVEL.Text + "','" + txtINDICE.Text + "','" + txtETOILE.Text + "', '" + txtPRFX.Text + "', '" + txtFAMILY.Text + "','" + txtMATRICULE.Text + "','" + txtOLL.Text + "');";

                bool isSuccess = DBAccess.ExecuteQuery(Query);
                Console.WriteLine(isSuccess + " / " + Query);
                if (isSuccess)
                {
                    MessageBox.Show("Item Imported Successfully", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    txtEPN.Enabled = false;
                    txtCPN1.Enabled = false;
                    txtCPN2.Enabled = false;
                    txtCPN3.Enabled = false;
                    txtALERT1.Enabled = false;
                    txtALERT2.Enabled = false;
                    txtALERT3.Enabled = false;
                    txtRELEASE.Enabled = false;
                    txtFCOSTUMER.Enabled = false;
                    txtFAMILY.Enabled = false;
                    txtMATRICULE.Enabled = false;
                    txtOLL.Enabled = false;
                    txtPRFX.Enabled = false;
                    txtLOT.Enabled = false;
                    txtLEVEL.Enabled = false;
                    txtQUANTITE.Enabled = false;
                    txtINDICE.Enabled = false;
                    txtETOILE.Enabled = false;


                    txtEPN.Text = "EPN"; 
                    txtCPN1.Text = "CPN 1"; 
                    txtCPN2.Text = "CPN 2";
                    txtCPN3.Text = "CPN 3";
                    txtALERT1.Text = "ALERT 1";
                    txtALERT2.Text = "ALERT 2"; 
                    txtALERT3.Text = "ALERT 3"; 
                    txtRELEASE.Text = "RELEASE"; 
                    txtFCOSTUMER.Text = "FIRST CUSTOMER"; 
                    txtFAMILY.Text = "FAMILY"; 
                    txtMATRICULE.Text = "OPERATOR ID"; 
                    txtLOT.Text = "LOT"; 
                    txtLEVEL.Text = "LEVEL"; 
                    txtINDICE.Text = "INDICE"; 
                    txtETOILE.Text = "ETOILE"; 
                    txtPRFX.Text = "PREFIX"; 
                    txtOLL.Text = "OLL";

                    dgvRef.Enabled = true;
                    btnNew.Text = "New";

                    
                    dgvRef.DataSource = DBAccess.ExecuteQuery2("SELECT [EPN] FROM[ETQSARCHIVE].[dbo].[REFERENCE]");

                }
                Console.WriteLine("update " + btnEdition.Text);
            }
            else
            {

                txtEPN.Enabled = true;
                txtCPN1.Enabled = true;
                txtCPN2.Enabled = true;
                txtCPN3.Enabled = true;
                txtALERT1.Enabled = true;
                txtALERT2.Enabled = true;
                txtALERT3.Enabled = true;
                txtRELEASE.Enabled = true;
                txtFCOSTUMER.Enabled = true;
                txtFAMILY.Enabled = true;
                txtMATRICULE.Enabled = true;
                txtOLL.Enabled = true;
                txtPRFX.Enabled = true;
                txtLOT.Enabled = true;
                txtLEVEL.Enabled = true;
                txtQUANTITE.Enabled = true;
                txtINDICE.Enabled = true;
                txtETOILE.Enabled = true;

                txtCPN1.ForeColor = Color.Black;
                txtCPN2.ForeColor = Color.Black;
                txtCPN3.ForeColor = Color.Black;
                txtALERT1.ForeColor = Color.Black;
                txtALERT2.ForeColor = Color.Black;
                txtALERT3.ForeColor = Color.Black;
                txtRELEASE.ForeColor = Color.Black;
                txtFCOSTUMER.ForeColor = Color.Black;
                txtFAMILY.ForeColor = Color.Black;
                txtMATRICULE.ForeColor = Color.Black;
                txtOLL.ForeColor = Color.Black;
                txtPRFX.ForeColor = Color.Black;
                txtLOT.ForeColor = Color.Black;
                txtLEVEL.ForeColor = Color.Black;
                txtQUANTITE.ForeColor = Color.Black;
                txtINDICE.ForeColor = Color.Black;
                txtETOILE.ForeColor = Color.Black;
                txtALERT3.ForeColor = Color.Black;
                

                txtEPN.Text = "EPN";
                txtCPN1.Text = "CPN 1";
                txtCPN2.Text = "CPN 2";
                txtCPN3.Text = "CPN 3";
                txtALERT1.Text = "ALERT 1";
                txtALERT2.Text = "ALERT 2";
                txtALERT3.Text = "ALERT 3";
                txtRELEASE.Text = "RELEASE";
                txtFCOSTUMER.Text = "FIRST CUSTOMER";
                txtFAMILY.Text = "FAMILY";
                txtMATRICULE.Text = "OPERATOR ID";
                txtLOT.Text = "LOT";
                txtLEVEL.Text = "LEVEL";
                txtINDICE.Text = "INDICE";
                txtETOILE.Text = "ETOILE";
                txtPRFX.Text = "PREFIX";
                txtOLL.Text = "OLL";


                dgvRef.Enabled = false;
                btnNew.Text = "INSERT";
                btnPrint.Enabled = false;
                btnEdition.Enabled = false;

            }

            txtCPN2.Focus();
            txtCPN3.Focus();
            txtALERT1.Focus();
            txtALERT2.Focus();
            txtALERT3.Focus();
            txtRELEASE.Focus();
            txtFCOSTUMER.Focus();
            txtFAMILY.Focus();
            txtMATRICULE.Focus();
            txtOLL.Focus();
            txtPRFX.Focus();
            txtLOT.Focus();
            txtLEVEL.Focus();
            txtQUANTITE.Focus();
            txtINDICE.Focus();
            txtETOILE.Focus();
            txtCPN1.Focus();
            txtEPN.Focus();
        }



        private void txtSearch_Enter(object sender, EventArgs e)
        {
            if (txtSearch.Text == "Search ...")
            {
                txtSearch.Text = "";
                txtSearch.ForeColor = Color.Black;
            }
        }

        private void txtSearch_Leave(object sender, EventArgs e)
        {
            if (txtSearch.Text == "")
            {
                txtSearch.Text = "Search ...";
                txtSearch.ForeColor = Color.Silver;
            }
        }

        private void guna2DataGridView1_Click(object sender, EventArgs e)
        {
            string Query = "";
            try
            {
                Query = "SELECT * FROM[ETQSARCHIVE].[dbo].[REFERENCE] WHERE [EPN] ='" + dgvRef.SelectedCells[0].Value + "'";
            }
            catch
            {
                return;
            }
            // = "SELECT * FROM[ETQSARCHIVE].[dbo].[REFERENCE] WHERE [EPN] ='" + dgvRef.SelectedCells[0].Value + "'";
            var Table = new DataTable();

            try
            {
                btnPrint.Enabled = true;
                btnEdition.Enabled = true;
                DBAccess.FillDataTable(Query, Table);


                string Q = "SELECT [NOM_ETI] FROM[ETQSARCHIVE].[dbo].[ETIQUETTE] WHERE [ID_ETIQUETTE] ='" + Table.Rows[0][18].ToString() + "'";
                var T = new DataTable();
                DBAccess.FillDataTable(Q, T);

                if (T.Rows.Count != 0)
                {
                    cbTemplate.SelectedIndex = cbTemplate.FindStringExact(T.Rows[0][0].ToString());
                }
                else
                {
                    MessageBox.Show("Template for this reference has been Deleted");
                }
                

                txtEPN.Text = Table.Rows[0][1].ToString();
                txtCPN1.Text = Table.Rows[0][2].ToString();
                txtCPN2.Text = Table.Rows[0][3].ToString();
                txtCPN3.Text = Table.Rows[0][4].ToString();
                txtALERT1.Text = Table.Rows[0][5].ToString();
                txtALERT2.Text = Table.Rows[0][6].ToString();
                txtALERT3.Text = Table.Rows[0][7].ToString();
                txtRELEASE.Text = Table.Rows[0][8].ToString();
                txtFCOSTUMER.Text = Table.Rows[0][9].ToString();

                txtLOT.Text = Table.Rows[0][10].ToString();
                txtLEVEL.Text = Table.Rows[0][11].ToString();
                txtINDICE.Text = Table.Rows[0][12].ToString();
                txtETOILE.Text = Table.Rows[0][13].ToString();
                txtPRFX.Text = Table.Rows[0][14].ToString();
                txtFAMILY.Text = Table.Rows[0][15].ToString();
                txtMATRICULE.Text = Table.Rows[0][16].ToString();
                txtOLL.Text = Table.Rows[0][17].ToString();
                txtQUANTITE.Text = Table.Rows[0][18].ToString();

            }
            catch (System.IO.IOException ex )
            {
                MessageBox.Show(ex.Message);

            }
            catch (System.IndexOutOfRangeException ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {

            
           t =  cbTemplate.Text;


            try
            {
                File.ReadAllText(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\LabelPrint\\Template\\" + cbTemplate.Text);
            }
            catch
            {
                MessageBox.Show("Please select a Template label");
                return;
            }

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
            }

            if (fileName != null)
            {

                var csvTable = new DataTable();
                using (var csvReader = new CsvReader(new StreamReader(System.IO.File.OpenRead(fileName)), true))
                {
                    csvTable.Load(csvReader);
                }

                String printer = "";

                int year = dtpicker.Value.Year;
                int month = dtpicker.Value.Month;
                int day = dtpicker.Value.Day;
                int hour = dtpicker.Value.Hour;
                int minute = dtpicker.Value.Minute;
                progressBar1.Maximum = csvTable.Rows.Count;
                progressBar1.Minimum = 0;
                progressBar1.Step = 1;


                Console.WriteLine("csvTable.Rows.Count" + csvTable.Rows.Count);
                progressBar1.Visible = true;
                for (int i = 0; i < csvTable.Rows.Count; i++)
                {
                    string str = File.ReadAllText(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\LabelPrint\\Template\\" + cbTemplate.Text);
                    str = str.Replace("@EPN", csvTable.Rows[i][0].ToString());
                    str = str.Replace("@CPN1", csvTable.Rows[i][1].ToString());
                    str = str.Replace("@CPN2", csvTable.Rows[i][2].ToString());
                    str = str.Replace("@CPN3", csvTable.Rows[i][3].ToString());
                    str = str.Replace("@ALERT1", csvTable.Rows[i][4].ToString());
                    str = str.Replace("@ALERT2", csvTable.Rows[i][5].ToString());
                    str = str.Replace("@ALERT3", csvTable.Rows[i][6].ToString());
                    str = str.Replace("@RELEASE", csvTable.Rows[i][7].ToString());
                    str = str.Replace("@FIRST_CUSTOMER", csvTable.Rows[i][8].ToString());
                    str = str.Replace("@FAMILLE", csvTable.Rows[i][9].ToString());
                    str = str.Replace("@OPERID", csvTable.Rows[i][10].ToString());
                    str = str.Replace("@OLL", csvTable.Rows[i][11].ToString());
                    str = str.Replace("@LOT", csvTable.Rows[i][12].ToString());
                    str = str.Replace("@LEVEL", csvTable.Rows[i][13].ToString());
                    str = str.Replace("@INDICE", csvTable.Rows[i][14].ToString());
                    str = str.Replace("@ETOILE", csvTable.Rows[i][15].ToString());
                    str = str.Replace("@PRFX", csvTable.Rows[i][16].ToString());
                    str = str.Replace("@QUANTITE", csvTable.Rows[i][17].ToString());
                    
                    str = str.Replace("@JJ", day.ToString());
                    str = str.Replace("@MM", month.ToString());
                    str = str.Replace("@YY", year.ToString());
                    str = str.Replace("@HOURS", hour.ToString());
                    

                    if (jjmmyyyy.Checked == true)
                    {
                        str = str.Replace("@DATETIME", dtpicker.Value.ToString("dd/MM/yyyy"));
                        str = str.Replace("@YY", year.ToString());
                    }
                    else
                    {
                        str = str.Replace("@DATETIME", dtpicker.Value.ToString("dd/MM/yy"));
                        str = str.Replace("@YY", year.ToString().Substring(2));
                    }

                    int copy = Convert.ToInt32(csvTable.Rows[i][17]);

                    System.Threading.Thread.Sleep(1500);

                    if (i == 0)
                    {

                        PrintDialog pd = new PrintDialog();

                        pd.PrinterSettings = new PrinterSettings();
                        if (DialogResult.OK == pd.ShowDialog(this))
                        {
                            for (int cp = 0; cp < copy; cp++)
                            {
                                str = str.Replace("@MINUTE", minute.ToString());
                                str = str.Replace("@SECOND", DateTime.Now.Second.ToString());
                                minute = minute + 2;

                                RawPrinterHelper.SendStringToPrinter(pd.PrinterSettings.PrinterName, str);
                            }
                            printer = pd.PrinterSettings.PrinterName;
                        }
                    }
                    else
                    {
                        for (int cp = 0; cp < copy; cp++)
                        {
                            str = str.Replace("@MINUTE", minute.ToString());
                            minute = minute + 2;
                            RawPrinterHelper.SendStringToPrinter(printer, str);
                        }

                    }
                    progressBar1.PerformStep();
                    File.Delete(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\LabelPrint\\tmp.prn");

                }
                System.Threading.Thread.Sleep(1500);

                //Do something with the file, for example read text from it
                DialogResult dialogResult = MessageBox.Show("Do you want to importe CSV data to archive ?", "Import CSV", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {

                    ImportCSV printCSV = new ImportCSV();
                    printCSV.dgvImportCSV.Rows.Clear();

                    //----------------------------------------------------------------------------------
                    try
                    {
                        int ImportedRecord = 0, inValidItem = 0;
                        string SourceURl = "";

                        if (fileName != "")
                        {
                            if (fileName.EndsWith(".csv"))
                            {
                                DataTable dtNew = new DataTable();
                                dtNew = GetDataTabletFromCSVFile(fileName);
                                if (Convert.ToString(dtNew.Columns[0]).ToUpper() != "EPN")
                                {
                                    MessageBox.Show("Invalid Items File");

                                    return;
                                }

                                SourceURl = fileName;
                                if (dtNew.Rows != null && dtNew.Rows.ToString() != String.Empty)
                                {
                                    printCSV.dgvImportCSV.DataSource = dtNew;
                                }
                                foreach (DataGridViewRow row in printCSV.dgvImportCSV.Rows)
                                {
                                    if (Convert.ToString(row.Cells["EPN"].Value) == "" || row.Cells["EPN"].Value == null
                                        || Convert.ToString(row.Cells["CPN1"].Value) == "" || row.Cells["CPN1"].Value == null
                                        || Convert.ToString(row.Cells["CPN2"].Value) == "" || row.Cells["CPN2"].Value == null
                                        || Convert.ToString(row.Cells["CPN3"].Value) == "" || row.Cells["CPN3"].Value == null)
                                    {
                                        row.DefaultCellStyle.BackColor = Color.Red;
                                        inValidItem += 1;
                                    }
                                    else
                                    {
                                        ImportedRecord += 1;
                                    }
                                }
                                if (printCSV.dgvImportCSV.Rows.Count == 0)
                                {

                                    MessageBox.Show("There is no data in this file", "GAUTAM POS", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                }
                            }
                            else
                            {
                                MessageBox.Show("Selected File is Invalid, Please Select valid csv file.", "GAUTAM POS", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Exception " + ex);
                        Console.WriteLine("here ----------------");
                    }
                    //----------------------------------------------------------------------------------
                    progressBar1.Visible = false;
                    printCSV.ShowDialog(this);
                }
                else if (dialogResult == DialogResult.No)
                {
                    return;
                }
                dgvRef.DataSource = DBAccess.ExecuteQuery2("SELECT [EPN] FROM[ETQSARCHIVE].[dbo].[REFERENCE]");
                path = fileName;
            }

        }


        private void guna2Button3_Click(object sender, EventArgs e)
        {

            if (btnNew.Text.Equals("INSERT"))
            {


                if (txtEPN.Text == "" || txtEPN.Text == null || txtEPN.Text == "EPN")
                {
                    MessageBox.Show("EPN cannot be empty ", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (txtEPN.Text == "EPN") { txtEPN.Text = ""; }
                if (txtCPN1.Text == "CPN 1") { txtCPN1.Text = ""; }
                if (txtCPN2.Text == "CPN 2") { txtCPN2.Text = ""; }
                if (txtCPN3.Text == "CPN 3") { txtCPN3.Text = ""; }
                if (txtALERT1.Text == "ALERT 1") { txtALERT1.Text = ""; }
                if (txtALERT2.Text == "ALERT 2") { txtALERT2.Text = ""; }
                if (txtALERT3.Text == "ALERT 3") { txtALERT3.Text = ""; }
                if (txtRELEASE.Text == "RELEASE") { txtRELEASE.Text = ""; }
                if (txtFCOSTUMER.Text == "FIRST CUSTOMER") { txtFCOSTUMER.Text = ""; }
                if (txtFAMILY.Text == "FAMILY") { txtFAMILY.Text = ""; }
                if (txtMATRICULE.Text == "OPERATOR ID") { txtMATRICULE.Text = ""; }
                if (txtLOT.Text == "LOT") { txtLOT.Text = ""; }
                if (txtLEVEL.Text == "LEVEL") { txtLEVEL.Text = ""; }
                if (txtINDICE.Text == "INDICE") { txtINDICE.Text = ""; }
                if (txtETOILE.Text == "ETOILE") { txtETOILE.Text = ""; }
                if (txtPRFX.Text == "PREFIX") { txtPRFX.Text = ""; }
                if (txtOLL.Text == "OLL") { txtOLL.Text = ""; }

                string Quer = "SELECT * FROM[ETQSARCHIVE].[dbo].[ETIQUETTE] WHERE [NOM_ETI] ='" + cbTemplate.Text + "'";
                var Table = new DataTable();

                DBAccess.FillDataTable(Quer, Table);
                

                string idEti;
                try
                {
                    idEti = Table.Rows[0][0].ToString();
                    
                }
                catch
                {
                    MessageBox.Show("Please selecte/add a Template ", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                 

                string Query = "INSERT INTO [dbo].[REFERENCE]  ([EPN] ,[CPN1] ,[CPN2] ,[CPN3]  ,[ALERT1] ,[ALERT2] ,[ALERT3] ,[RELEASE] ,[CUSTOMER] ,[LOT] ,[LEVEL_] ,[INDICE] ,[ETOILE] ,[PRFX] ,[ID_FAMILY] ,[OPERATOR] ,[OLL], [ID_ETIQUETTE]) VALUES ('" + txtEPN.Text + "' , '" + txtCPN1.Text + "','" + txtCPN2.Text + "', '" + txtCPN3.Text + "','" + txtALERT1.Text + "', '" + txtALERT2.Text + "','" + txtALERT3.Text + "','" + txtRELEASE.Text + "','" + txtFCOSTUMER.Text + "','" + txtLOT.Text + "','" + txtLEVEL.Text + "','" + txtINDICE.Text + "','" + txtETOILE.Text + "', '" + txtPRFX.Text + "', '" + txtFAMILY.Text + "','" + txtMATRICULE.Text + "','" + txtOLL.Text + "','" + idEti + "');";
                bool isSuccess = DBAccess.ExecuteQuery(Query);
                Console.WriteLine(isSuccess + " / " + Query);
                if (isSuccess)
                {
                    MessageBox.Show("Item Imported Successfully", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    txtEPN.Enabled = false;
                    txtCPN1.Enabled = false;
                    txtCPN2.Enabled = false;
                    txtCPN3.Enabled = false;
                    txtALERT1.Enabled = false;
                    txtALERT2.Enabled = false;
                    txtALERT3.Enabled = false;
                    txtRELEASE.Enabled = false;
                    txtFCOSTUMER.Enabled = false;
                    txtFAMILY.Enabled = false;
                    txtMATRICULE.Enabled = false;
                    txtOLL.Enabled = false;
                    txtPRFX.Enabled = false;
                    txtLOT.Enabled = false;
                    txtLEVEL.Enabled = false;
                    txtQUANTITE.Enabled = false;
                    txtINDICE.Enabled = false;
                    txtETOILE.Enabled = false;


                    txtEPN.Text = "EPN";
                    txtCPN1.Text = "CPN 1";
                    txtCPN2.Text = "CPN 2";
                    txtCPN3.Text = "CPN 3";
                    txtALERT1.Text = "ALERT 1";
                    txtALERT2.Text = "ALERT 2";
                    txtALERT3.Text = "ALERT 3";
                    txtRELEASE.Text = "RELEASE";
                    txtFCOSTUMER.Text = "FIRST CUSTOMER";
                    txtFAMILY.Text = "FAMILY";
                    txtMATRICULE.Text = "OPERATOR ID";
                    txtLOT.Text = "LOT";
                    txtLEVEL.Text = "LEVEL";
                    txtINDICE.Text = "INDICE";
                    txtETOILE.Text = "ETOILE";
                    txtPRFX.Text = "PREFIX";
                    txtOLL.Text = "OLL";

                    dgvRef.Enabled = true;
                    btnNew.Text = "New";


                    dgvRef.DataSource = DBAccess.ExecuteQuery2("SELECT [EPN] FROM[ETQSARCHIVE].[dbo].[REFERENCE]");

                }
                Console.WriteLine("update " + btnEdition.Text);
            }
            else
            {

                txtEPN.Enabled = true;
                txtCPN1.Enabled = true;
                txtCPN2.Enabled = true;
                txtCPN3.Enabled = true;
                txtALERT1.Enabled = true;
                txtALERT2.Enabled = true;
                txtALERT3.Enabled = true;
                txtRELEASE.Enabled = true;
                txtFCOSTUMER.Enabled = true;
                txtFAMILY.Enabled = true;
                txtMATRICULE.Enabled = true;
                txtOLL.Enabled = true;
                txtPRFX.Enabled = true;
                txtLOT.Enabled = true;
                txtLEVEL.Enabled = true;
                txtQUANTITE.Enabled = true;
                txtINDICE.Enabled = true;
                txtETOILE.Enabled = true;

                txtCPN1.ForeColor = Color.Black;
                txtCPN2.ForeColor = Color.Black;
                txtCPN3.ForeColor = Color.Black;
                txtALERT1.ForeColor = Color.Black;
                txtALERT2.ForeColor = Color.Black;
                txtALERT3.ForeColor = Color.Black;
                txtRELEASE.ForeColor = Color.Black;
                txtFCOSTUMER.ForeColor = Color.Black;
                txtFAMILY.ForeColor = Color.Black;
                txtMATRICULE.ForeColor = Color.Black;
                txtOLL.ForeColor = Color.Black;
                txtPRFX.ForeColor = Color.Black;
                txtLOT.ForeColor = Color.Black;
                txtLEVEL.ForeColor = Color.Black;
                txtQUANTITE.ForeColor = Color.Black;
                txtINDICE.ForeColor = Color.Black;
                txtETOILE.ForeColor = Color.Black;
                txtALERT3.ForeColor = Color.Black;

                txtEPN.Text = "EPN";
                txtCPN1.Text = "CPN 1";
                txtCPN2.Text = "CPN 2";
                txtCPN3.Text = "CPN 3";
                txtALERT1.Text = "ALERT 1";
                txtALERT2.Text = "ALERT 2";
                txtALERT3.Text = "ALERT 3";
                txtRELEASE.Text = "RELEASE";
                txtFCOSTUMER.Text = "FIRST CUSTOMER";
                txtFAMILY.Text = "FAMILY";
                txtMATRICULE.Text = "OPERATOR ID";
                txtLOT.Text = "LOT";
                txtLEVEL.Text = "LEVEL";
                txtINDICE.Text = "INDICE";
                txtETOILE.Text = "ETOILE";
                txtPRFX.Text = "PREFIX";
                txtOLL.Text = "OLL";

                dgvRef.Enabled = false;
                btnNew.Text = "INSERT";
                btnPrint.Enabled = false;
                btnEdition.Enabled = false;
            }

            txtCPN2.Focus();
            txtCPN3.Focus();
            txtALERT1.Focus();
            txtALERT2.Focus();
            txtALERT3.Focus();
            txtRELEASE.Focus();
            txtFCOSTUMER.Focus();
            txtFAMILY.Focus();
            txtMATRICULE.Focus();
            txtOLL.Focus();
            txtPRFX.Focus();
            txtLOT.Focus();
            txtLEVEL.Focus();
            txtQUANTITE.Focus();
            txtINDICE.Focus();
            txtETOILE.Focus();
            txtCPN1.Focus();
            txtEPN.Focus();
        }

        private void guna2Button4_Click(object sender, EventArgs e)
        {

            if (txtEPN.Text == "" || txtEPN.Text == null || txtEPN.Text == "EPN")
            {
                return;
            }

            string Quer = "SELECT * FROM[ETQSARCHIVE].[dbo].[ETIQUETTE] WHERE [NOM_ETI] ='" + cbTemplate.Text + "'";
            var Table = new DataTable();

            DBAccess.FillDataTable(Quer, Table);
            string idEti = Table.Rows[0][0].ToString();

            if (btnEdition.Text.Equals("UPDATE"))
            {

                

                if (txtEPN.Text == "EPN") { txtEPN.Text = ""; }
                if (txtCPN1.Text == "CPN 1") { txtCPN1.Text = ""; }
                if (txtCPN2.Text == "CPN 2") { txtCPN2.Text = ""; }
                if (txtCPN3.Text == "CPN 3") { txtCPN3.Text = ""; }
                if (txtALERT1.Text == "ALERT 1") { txtALERT1.Text = ""; }
                if (txtALERT2.Text == "ALERT 2") { txtALERT2.Text = ""; }
                if (txtALERT3.Text == "ALERT 3") { txtALERT3.Text = ""; }
                if (txtRELEASE.Text == "RELEASE") { txtRELEASE.Text = ""; }
                if (txtFCOSTUMER.Text == "FIRST CUSTOMER") { txtFCOSTUMER.Text = ""; }
                if (txtFAMILY.Text == "FAMILY") { txtFAMILY.Text = ""; }
                if (txtMATRICULE.Text == "OPERATOR ID") { txtMATRICULE.Text = ""; }
                if (txtLOT.Text == "LOT") { txtLOT.Text = ""; }
                if (txtLEVEL.Text == "LEVEL") { txtLEVEL.Text = ""; }
                if (txtINDICE.Text == "INDICE") { txtINDICE.Text = ""; }
                if (txtETOILE.Text == "ETOILE") { txtETOILE.Text = ""; }
                if (txtPRFX.Text == "PREFIX") { txtPRFX.Text = ""; }
                if (txtOLL.Text == "OLL") { txtOLL.Text = ""; }

                string Query = "UPDATE [dbo].[REFERENCE] SET [CPN1] = '" + txtCPN1.Text + "',[CPN2] = '" + txtCPN2.Text + "',[CPN3] = '" + txtCPN3.Text + "' ,[ALERT1] = '" + txtALERT1.Text + "',[ALERT2] = '" + txtALERT2.Text + "',[ALERT3] = '" + txtALERT3.Text + "',[RELEASE] ='" + txtRELEASE.Text + "',[CUSTOMER] = '" + txtFCOSTUMER.Text + "',[LOT] = '" + txtLOT.Text + "',[LEVEL_] = '" + txtLEVEL.Text + "' ,[INDICE] = '" + txtINDICE.Text + "' ,[ETOILE] = '" + txtETOILE.Text + "' ,[PRFX] = '" + txtPRFX.Text + "' ,[ID_FAMILY] = '" + txtFAMILY.Text + "' ,[OPERATOR] = '" + txtMATRICULE.Text + "' ,[OLL] = '" + txtOLL.Text + "' ,[ID_ETIQUETTE] = '"+idEti+"' WHERE [EPN] = '" + txtEPN.Text + "'";

                bool isSuccess = DBAccess.ExecuteQuery(Query);
                Console.WriteLine(isSuccess + " / " + Query);
                if (isSuccess)
                {

                    MessageBox.Show("Item Updated Successfully", "", MessageBoxButtons.OK, MessageBoxIcon.Information);


                    txtCPN1.Enabled = false;
                    txtCPN2.Enabled = false;
                    txtCPN3.Enabled = false;
                    txtALERT1.Enabled = false;
                    txtALERT2.Enabled = false;
                    txtALERT3.Enabled = false;
                    txtRELEASE.Enabled = false;
                    txtFCOSTUMER.Enabled = false;
                    txtFAMILY.Enabled = false;
                    txtMATRICULE.Enabled = false;
                    txtOLL.Enabled = false;
                    txtPRFX.Enabled = false;
                    txtLOT.Enabled = false;
                    txtLEVEL.Enabled = false;
                    txtQUANTITE.Enabled = false;
                    txtINDICE.Enabled = false;
                    txtETOILE.Enabled = false;


                    txtEPN.Text = "EPN";
                    txtCPN1.Text = "CPN 1";
                    txtCPN2.Text = "CPN 2";
                    txtCPN3.Text = "CPN 3";
                    txtALERT1.Text = "ALERT 1";
                    txtALERT2.Text = "ALERT 2";
                    txtALERT3.Text = "ALERT 3";
                    txtRELEASE.Text = "RELEASE";
                    txtFCOSTUMER.Text = "FIRST CUSTOMER";
                    txtFAMILY.Text = "FAMILY";
                    txtMATRICULE.Text = "OPERATOR ID";
                    txtLOT.Text = "LOT";
                    txtLEVEL.Text = "LEVEL";
                    txtINDICE.Text = "INDICE";
                    txtETOILE.Text = "ETOILE";
                    txtPRFX.Text = "PREFIX";
                    txtOLL.Text = "OLL";

                    btnEdition.Text = "EDITE";
                    btnNew.Enabled = true;
                    btnPrint.Enabled = false;
                    btnEdition.Enabled = false;
                    btnNew.Enabled = true;
                    dgvRef.Enabled = true;

                }

            }
            else
            {
                btnNew.Enabled = false;
                btnPrint.Enabled = false;
                dgvRef.Enabled = false;

                Console.WriteLine("edite " + btnEdition.Text);
                // txtEPN.Enabled = true;
                txtCPN1.Enabled = true;
                txtCPN2.Enabled = true;
                txtCPN3.Enabled = true;
                txtALERT1.Enabled = true;
                txtALERT2.Enabled = true;
                txtALERT3.Enabled = true;
                txtRELEASE.Enabled = true;
                txtFCOSTUMER.Enabled = true;
                txtFAMILY.Enabled = true;
                txtMATRICULE.Enabled = true;
                txtOLL.Enabled = true;
                txtPRFX.Enabled = true;
                txtLOT.Enabled = true;
                txtLEVEL.Enabled = true;
                txtQUANTITE.Enabled = true;
                txtINDICE.Enabled = true;
                txtETOILE.Enabled = true;

                txtCPN1.ForeColor = Color.Black;
                txtCPN2.ForeColor = Color.Black;
                txtCPN3.ForeColor = Color.Black;
                txtALERT1.ForeColor = Color.Black;
                txtALERT2.ForeColor = Color.Black;
                txtALERT3.ForeColor = Color.Black;
                txtRELEASE.ForeColor = Color.Black;
                txtFCOSTUMER.ForeColor = Color.Black;
                txtFAMILY.ForeColor = Color.Black;
                txtMATRICULE.ForeColor = Color.Black;
                txtOLL.ForeColor = Color.Black;
                txtPRFX.ForeColor = Color.Black;
                txtLOT.ForeColor = Color.Black;
                txtLEVEL.ForeColor = Color.Black;
                txtQUANTITE.ForeColor = Color.Black;
                txtINDICE.ForeColor = Color.Black;
                txtETOILE.ForeColor = Color.Black;
                txtALERT3.ForeColor = Color.Black;
                btnEdition.Text = "UPDATE";

            }
            txtEPN.Focus();
            txtCPN2.Focus();
            txtCPN3.Focus();
            txtALERT1.Focus();
            txtALERT2.Focus();
            txtALERT3.Focus();
            txtRELEASE.Focus();
            txtFCOSTUMER.Focus();
            txtFAMILY.Focus();
            txtMATRICULE.Focus();
            txtOLL.Focus();
            txtPRFX.Focus();
            txtLOT.Focus();
            txtLEVEL.Focus();
            txtQUANTITE.Focus();
            txtINDICE.Focus();
            txtETOILE.Focus();
            txtCPN1.Focus();
        }

        private void guna2Button2_Click_1(object sender, EventArgs e)
        {

            AddTemplate addTemplate = new AddTemplate();
            addTemplate.ShowDialog(this);


        }

        private void guna2ComboBox1_DropDown(object sender, EventArgs e)
        {
            DataTable table = new DataTable();

            DBAccess.FillDataTable("SELECT [NOM_ETI] FROM [ETQSARCHIVE].[dbo].[ETIQUETTE]", table);

            
            //Console.WriteLine("test "+table.Rows[0][0].ToString());

            cbTemplate.DataSource = table;
            //cbTemplate.DisplayMember = "File Name";
            cbTemplate.ValueMember = "NOM_ETI"; 



            /*
            string path = Application.StartupPath + "/Template";
            DataTable table = new DataTable();
            table.Columns.Add("File Name");
            table.Columns.Add("File Path");

            string[] files = Directory.GetFiles(path);

            for (int i = 0; i < files.Length; i++)
            {
                FileInfo file = new FileInfo(files[i]);
                table.Rows.Add(file.Name, path + "\\" + file.Name);
            }

            cbTemplate.DataSource = table;
            cbTemplate.DisplayMember = "File Name";
            cbTemplate.ValueMember = "File Path";*/
        }


        //work
        private void btnPrint_Click_1(object sender, EventArgs e)
        {
            
            int year = dtpicker.Value.Year;
            int month = dtpicker.Value.Month;
            int day = dtpicker.Value.Day;
            int minute = dtpicker.Value.Minute;
            int hour = dtpicker.Value.Hour;

            try
            {
                File.ReadAllText(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\LabelPrint\\Template\\" + cbTemplate.Text);
            }
            catch
            {
                MessageBox.Show("Please select a Template label");
                return;
            }

            int copy = (int)numericUpDown1.Value;
            System.Threading.Thread.Sleep(1500);
            PrintDialog pd = new PrintDialog();

            if (DialogResult.OK == pd.ShowDialog(this))
            {

                for (int cp = 0; cp < copy; cp++)
                {

                    string str = File.ReadAllText(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\LabelPrint\\Template\\" + cbTemplate.Text);

                    str = str.Replace("@EPN", txtEPN.Text);
                    str = str.Replace("@CPN1", txtCPN1.Text);
                    str = str.Replace("@CPN2", txtCPN2.Text);
                    str = str.Replace("@CPN3", txtCPN3.Text);
                    str = str.Replace("@ALERT1", txtALERT1.Text);
                    str = str.Replace("@ALERT2", txtALERT2.Text);
                    str = str.Replace("@ALERT3", txtALERT3.Text);
                    str = str.Replace("@RELEASE", txtRELEASE.Text);
                    str = str.Replace("@FIRST_CUSTOMER", txtFCOSTUMER.Text);
                    str = str.Replace("@FAMILLE", txtFAMILY.Text);
                    str = str.Replace("@OPERID", txtMATRICULE.Text);
                    str = str.Replace("@OLL", txtOLL.Text);
                    str = str.Replace("@LOT", txtLOT.Text);
                    str = str.Replace("@LEVEL", txtLEVEL.Text);
                    str = str.Replace("@INDICE", txtINDICE.Text);
                    str = str.Replace("@ETOILE", txtETOILE.Text);
                    str = str.Replace("@PRFX", txtPRFX.Text);
                    str = str.Replace("@HOURS", dtpicker.Value.ToString("hh:mm"));
                    Random rnd = new Random();
                    rnd.Next(1,60);
                    str = str.Replace("@COUNTER", rnd.Next(10, 60) + rnd.Next(10, 60).ToString());
                  
                    //minute += 2;

                    str = str.Replace("@JJ", day.ToString());
                    str = str.Replace("@MM", month.ToString());




                    if (jjmmyyyy.Checked == true)
                    {
                        str = str.Replace("@DATETIME", dtpicker.Value.ToString("dd/MM/yyyy"));
                        str = str.Replace("@YY", year.ToString());
                    }
                    else
                    {
                        str = str.Replace("@DATETIME", dtpicker.Value.ToString("dd/MM/yy"));
                        str = str.Replace("@YY", year.ToString().Substring(2));
                    }

                    RawPrinterHelper.SendStringToPrinter(pd.PrinterSettings.PrinterName, str);


                    dtpicker.Value = dtpicker.Value.AddMinutes((int)numericUpDown2.Value);

                }
            }
            File.Delete(Application.StartupPath + "/tmp.prn");
        }

        private void btnPrintCSV_Click(object sender, EventArgs e)
        {

            ImportCSV importCSV = new ImportCSV();
            importCSV.ShowDialog(this);

        }

        private void guna2Button3_Click_1(object sender, EventArgs e)
        {
            if (cbTemplate.Text.Length != 0 )
            {
                DialogResult dialogResult = MessageBox.Show("Do you want to delete Template ?", "Delete Template", MessageBoxButtons.YesNo);

                if (dialogResult == DialogResult.No)
                {
                    return;
                }
                bool isSuccessDeletedBD = DBAccess.ExecuteQuery("DELETE FROM [dbo].[ETIQUETTE] WHERE [NOM_ETI] = ('" + cbTemplate.Text + "')");
                File.Delete(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\LabelPrint\\Template\\" + cbTemplate.Text);

                if (isSuccessDeletedBD)
                {
                    MessageBox.Show("Template Deleted Successfully", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }
    }
                    
                   
}
    
