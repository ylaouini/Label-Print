using System;
using System.Data;
using System.Drawing;
using System.IO;

using System.Windows.Forms;

namespace Label_Print
{
    public partial class AddTemplate : Form
    {
        public AddTemplate()
        {
            InitializeComponent();
        }
        private string fileName = null;

        private void btnChooseTemplate_Click(object sender, EventArgs e)
        {
            
            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                //openFileDialog1.InitialDirectory = path;
                openFileDialog1.Filter = "csv files(*.*) | *.*";
                openFileDialog1.FilterIndex = 2;
                openFileDialog1.RestoreDirectory = true;

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    fileName = openFileDialog1.FileName;
                    Console.WriteLine("path: " + openFileDialog1.Title);
                    string[] s = (openFileDialog1.FileName.ToString()).Split('\\');
                    int count = s.Length; ;
                    guna2TextBox1.Text = s[count-1];
                }
            }

            if (fileName != null)
            {

                lblEPN.Text = "✘ @EPN"; lblEPN.ForeColor = Color.Black;
                lblCPN1.Text = "✘ @CPN1"; lblCPN1.ForeColor = Color.Black;
                lblCPN2.Text = "✘ @CPN2"; lblCPN2.ForeColor = Color.Black;
                lblCPN3.Text = "✘ @CPN3"; lblCPN3.ForeColor = Color.Black;
                lblALERT1.Text = "✘ @ALERT1"; lblALERT1.ForeColor = Color.Black;
                lblALERT2.Text = "✘ @ALERT2"; lblALERT2.ForeColor = Color.Black;
                lblALERT3.Text = "✘ @ALERT3"; lblALERT3.ForeColor = Color.Black;
                lblRELEASE.Text = "✘ @RELEASE"; lblRELEASE.ForeColor = Color.Black;
                lblFIRST_CUSTOMER.Text = "✘ @FIRST_CUSTOMER"; lblFIRST_CUSTOMER.ForeColor = Color.Black;
                lblFAMILLE.Text = "✘ @FAMILLE"; lblFAMILLE.ForeColor = Color.Black;
                lblOPERID.Text = "✘ @OPERID"; lblOPERID.ForeColor = Color.Black;
                lblOLL.Text = "✘ @OLL"; lblOLL.ForeColor = Color.Black;
                lblLOT.Text = "✘ @LOT"; lblLOT.ForeColor = Color.Black;
                lblLEVEL.Text = "✘ @LEVEL"; lblLEVEL.ForeColor = Color.Black;
                lblINDICE.Text = "✘ @INDICE"; lblINDICE.ForeColor = Color.Black;
                lblETOILE.Text = "✘ @ETOILE"; lblETOILE.ForeColor = Color.Black;
                lblPRFX.Text = "✘ @PRFX"; lblPRFX.ForeColor = Color.Black;
                lblHOURS.Text = "✘ @HOURS"; lblHOURS.ForeColor = Color.Black;
                lblYY.Text = "✘ @YY"; lblYY.ForeColor = Color.Black;
                lblJJ.Text = "✘ @JJ"; lblJJ.ForeColor = Color.Black;
                lblMM.Text = "✘ @MM"; lblMM.ForeColor = Color.Black;
                lblDATETIME.Text = "✘ @DATETIME"; lblDATETIME.ForeColor = Color.Black;
                lblMINUTE.Text = "✘ @MINUTE"; lblMINUTE.ForeColor = Color.Black;
                lblCOUNTER.Text = "✘ @COUNTER"; lblCOUNTER.ForeColor = Color.Black;


                string str = File.ReadAllText(fileName);
                if (str.Contains("@EPN")) { lblEPN.Text = "✔ @EPN"; lblEPN.ForeColor = Color.Green; }
                if (str.Contains("@CPN1")) { lblCPN1.Text = "✔ @CPN1"; lblCPN1.ForeColor = Color.Green; }
                if (str.Contains("@CPN2")) { lblCPN2.Text = "✔ @CPN2"; lblCPN2.ForeColor = Color.Green; }
                if (str.Contains("@CPN3")) { lblCPN3.Text = "✔ @CPN3"; lblCPN3.ForeColor = Color.Green; }
                if (str.Contains("@ALERT1")) { lblALERT1.Text = "✔ @ALERT1"; lblALERT1.ForeColor = Color.Green; }
                if (str.Contains("@ALERT2")) { lblALERT2.Text = "✔ @ALERT2"; lblALERT2.ForeColor = Color.Green; }
                if (str.Contains("@ALERT3")) { lblALERT3.Text = "✔ @ALERT3"; lblALERT3.ForeColor = Color.Green; }
                if (str.Contains("@RELEASE")) { lblRELEASE.Text = "✔ @RELEASE"; lblRELEASE.ForeColor = Color.Green; }
                if (str.Contains("@FIRST_CUSTOMER")) { lblFIRST_CUSTOMER.Text = "✔ @FIRST_CUSTOMER"; lblFIRST_CUSTOMER.ForeColor = Color.Green; }
                if (str.Contains("@FAMILLE")) { lblFAMILLE.Text = "✔ @FAMILLE"; lblFAMILLE.ForeColor = Color.Green; }
                if (str.Contains("@OPERID")) { lblOPERID.Text = "✔ @OPERID"; lblOPERID.ForeColor = Color.Green; }
                if (str.Contains("@OLL")) { lblOLL.Text = "✔ @OLL"; lblOLL.ForeColor = Color.Green; }
                if (str.Contains("@LOT")) { lblLOT.Text = "✔ @LOT"; lblLOT.ForeColor = Color.Green; }
                if (str.Contains("@LEVEL")) { lblLEVEL.Text = "✔ @LEVEL"; lblLEVEL.ForeColor = Color.Green; }
                if (str.Contains("@INDICE")) { lblINDICE.Text = "✔ @INDICE"; lblINDICE.ForeColor = Color.Green; }
                if (str.Contains("@ETOILE")) { lblETOILE.Text = "✔ @ETOILE"; lblETOILE.ForeColor = Color.Green; }
                if (str.Contains("@PRFX")) { lblPRFX.Text = "✔ @PRFX"; lblPRFX.ForeColor = Color.Green; }
                if (str.Contains("@HOURS")) { lblHOURS.Text = "✔ @HOURS"; lblHOURS.ForeColor = Color.Green; }
                if (str.Contains("@YY")) { lblYY.Text = "✔ @YY"; lblYY.ForeColor = Color.Green; }
                if (str.Contains("@JJ")) { lblJJ.Text = "✔ @JJ"; lblJJ.ForeColor = Color.Green; }
                if (str.Contains("@MM")) { lblMM.Text = "✔ @MM"; lblMM.ForeColor = Color.Green; }
                if (str.Contains("@DATETIME")) { lblDATETIME.Text = "✔ @DATETIME"; lblDATETIME.ForeColor = Color.Green; }
                if (str.Contains("@MINUTE")) { lblMINUTE.Text = "✔ @MINUTE"; lblMINUTE.ForeColor = Color.Green; }
                if (str.Contains("@COUNTER")) { lblCOUNTER.Text = "✔ @COUNTER"; lblCOUNTER.ForeColor = Color.Green; }

                if (fileName != "" || fileName != null || guna2TextBox1.Text != "" || guna2TextBox1.Text != null)
                {
                    btnImportTemplate.Enabled = true;
                }
            }
        }

        private void btnImportTemplate_Click(object sender, EventArgs e)
        {
            if (lblEPN.Text == "✘ @EPN" || lblCPN1.Text == "✘ @CPN1" || lblCPN2.Text == "✘ @CPN2" || lblCPN3.Text == "✘ @CPN3" || lblALERT1.Text == "✘ @ALERT1" || lblALERT2.Text == "✘ @ALERT2" || lblALERT3.Text == "✘ @ALERT3" || lblRELEASE.Text == "✘ @RELEASE" || lblFIRST_CUSTOMER.Text == "✘ @FIRST_CUSTOMER" || lblFAMILLE.Text == "✘ @FAMILLE" || lblOPERID.Text == "✘ @OPERID" || lblOLL.Text == "✘ @OLL" || lblLOT.Text == "✘ @LOT" || lblLEVEL.Text == "✘ @LEVEL" || lblINDICE.Text == "✘ @INDICE" || lblETOILE.Text == "✘ @ETOILE" || lblPRFX.Text == "✘ @PRFX" || lblHOURS.Text == "✘ @HOURS" || lblYY.Text == "✘ @YY" || lblJJ.Text == "✘ @JJ" || lblMM.Text == "✘ @MM" || lblDATETIME.Text == "✘ @DATETIME"  || lblCOUNTER.Text == "✘ @COUNTER")
            {
                DialogResult dialogResult = MessageBox.Show("you have missing variable(s) in your template, if you import this template you risk to have missing variable in your label ! \nDo you want to import ?", "Import CSV", MessageBoxButtons.YesNo);

                if (dialogResult == DialogResult.No)
                {
                    return;
                }
            }

            bool isSuccessTocopy;
            try
            {
                File.Copy(fileName, Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\LabelPrint\\Template\\" + guna2TextBox1.Text);
                isSuccessTocopy = true;
            }
            catch (System.IO.IOException ex)
            {
                MessageBox.Show(ex.Message + "\n You can chose another name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Console.WriteLine();
                isSuccessTocopy = false;
                return;
            }


            bool isSuccessToBD = DBAccess.ExecuteQuery("INSERT INTO [dbo].[ETIQUETTE] ([NOM_ETI]) VALUES ('" + guna2TextBox1.Text + "')");

            if (isSuccessToBD && isSuccessTocopy)
            {
                MessageBox.Show("Template Imported Successfully", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                btnImportTemplate.Enabled = false;
            }
            else
            {
                File.Delete(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\LabelPrint\\Template\\" + guna2TextBox1.Text);
            }


        }


        }
    }

