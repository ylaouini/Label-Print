using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
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
        public static String[,] array = new String[2000, 5];
        public static int ligne = 0;
        private void btnPrint_Click(object sender, EventArgs e)
        {
            String printer = "";
            for (int line = 0; line < ligne; line++)
            {
                string str = File.ReadAllText(Application.StartupPath + "/label.prn");
                str = str.Replace("@xcode", array[line, 0]);
                str = str.Replace("@epn", array[line, 1]);
                str = str.Replace("@cordX", array[line, 2]);
                str = str.Replace("@cordY", array[line, 3]);
                System.Threading.Thread.Sleep(1500);
                //File.WriteAllText(Application.StartupPath + "/tmp.prn", str);
                //System.Diagnostics.Process.Start("cmd.exe", "/c print " + Application.StartupPath + "\\tmp.prn");
                if (line == 0)
                {
                    PrintDialog pd = new PrintDialog();
                    pd.PrinterSettings = new PrinterSettings();
                    if (DialogResult.OK == pd.ShowDialog(this))
                    {
                      
                        RawPrinterHelper.SendStringToPrinter(pd.PrinterSettings.PrinterName, str);
                        printer = pd.PrinterSettings.PrinterName;
                    }

                }
                else
                {
                    RawPrinterHelper.SendStringToPrinter(printer, str);
                }
                File.Delete(Application.StartupPath + "/tmp.prn");

            }
        }

        private void btnCSVSelect_Click(object sender, EventArgs e)
        {
            string fileName = null;
            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                openFileDialog1.InitialDirectory = path;
                openFileDialog1.Filter = "csv files (*.csv)|*.csv";
                openFileDialog1.FilterIndex = 2;
                openFileDialog1.RestoreDirectory = true;

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    fileName = openFileDialog1.FileName;
                }
            }

            if (fileName != null)
            {
                //Do something with the file, for example read text from it
                path = fileName;
                textBox1.Text = path;
            }
            using (var reader = new StreamReader(path))
            {

                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(';');
                    array[ligne, 0] = values[0];
                    array[ligne, 1] = values[1];
                    array[ligne, 2] = values[2];
                    array[ligne, 3] = values[3];
                    ligne++;
                }
            }
        }
    }
