using System.Data.SqlClient;
using System;
using System.Data;
using System.Windows.Forms;

namespace Label_Print
{
    class DBAccess
    {
        private static SqlConnection objConnection;
        private static SqlDataAdapter objDataAdapter;

        private static void OpenConnection()
        {
            try
            {
                if (objConnection == null)
                {
                    objConnection = new SqlConnection(@"Data Source = (localdb)\v11.0; Initial Catalog = ETQSARCHIVE; AttachDbFilename=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\LabelPrint\\BD\\ETQSARCHIVE.mdf;Database=ETQSARCHIVE ; Trusted_Connection = True; ");
                    objConnection.Open();
                    //Console.Write("OpenConnection1");
                }
                else
                {
                    if (objConnection.State != ConnectionState.Open)
                    {
                        objConnection = new SqlConnection(@"Data Source = (localdb)\v11.0; Initial Catalog = ETQSARCHIVE; AttachDbFilename=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\LabelPrint\\BD\\ETQSARCHIVE.mdf;Database=ETQSARCHIVE ; Trusted_Connection = True; ");
                        objConnection.Open();
                        //Console.Write("OpenConnection2");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message ,"", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Console.Write("Opppps"+ex);
            }
        }

        private static void CloseConnection()
        {
            try
            {
                if (!(objConnection == null))
                {
                    if (objConnection.State == ConnectionState.Open)
                    {
                        objConnection.Close();
                        objConnection.Dispose();
                    }
                }
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
        }

        public static DataTable FillDataTable(string Query, DataTable Table)
        {

            OpenConnection();
            try
            {
                objDataAdapter = new SqlDataAdapter(Query, objConnection);
                objDataAdapter.Fill(Table);
                objDataAdapter.Dispose();
                CloseConnection();


                return Table;

            }
            catch
            {
                return null;
            }
        }


        public static SqlDataReader ExecuteReader(string cmd)
        {

            try
            {
                SqlDataReader objReader;
                //objConnection = new SqlConnection("Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + Application.StartupPath + "/BD/ETQSARCHIVE.mdb;");
                objConnection = new SqlConnection(@"Data Source = (localdb)\v11.0; Initial Catalog = ETQSARCHIVE; AttachDbFilename=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\LabelPrint\\BD\\ETQSARCHIVE.mdf;Database=ETQSARCHIVE ; Trusted_Connection = True; ");
                OpenConnection();
                SqlCommand cmdRedr = new SqlCommand(cmd, objConnection);
                objReader = cmdRedr.ExecuteReader(CommandBehavior.CloseConnection);
                cmdRedr.Dispose();
                return objReader;
            }
            catch
            {
                return null;
            }
        }
        public static bool ExecuteQuery(string query)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(@"Data Source = (localdb)\v11.0; Initial Catalog = ETQSARCHIVE; AttachDbFilename=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\LabelPrint\\BD\\ETQSARCHIVE.mdf;Database=ETQSARCHIVE ; Trusted_Connection = True; "))
                {
                    using (SqlCommand cmd = new SqlCommand(query, connection))
                    {
                        connection.Open();
                        cmd.ExecuteNonQuery();
                        connection.Close();
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                string s = "";
                s = ex.Message.Replace("Violation of PRIMARY KEY constraint 'PK_REFERENCE'. Cannot insert duplicate key in object 'dbo.REFERENCE'. The duplicate key value is", "").Replace("The statement has been terminated.", "");
               
                Console.WriteLine("Cannot insert duplicate EPN !. \n The duplicate EPN is :\n" + s);
                MessageBox.Show("Cannot insert duplicate EPN!. \n EPN already exist :\n" + s, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        public static DataTable ExecuteQuery2(string query)
        {
            DataTable dataTable = new DataTable();
            try
            {
                using (SqlConnection connection = new SqlConnection(@"Data Source = (localdb)\v11.0; Initial Catalog = ETQSARCHIVE; AttachDbFilename=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\LabelPrint\\BD\\ETQSARCHIVE.mdf;Database=ETQSARCHIVE ; Trusted_Connection = True; "))
                {
                    using (SqlCommand cmd = new SqlCommand(query, connection))
                    {
                        connection.Open();
                        SqlDataReader reader = cmd.ExecuteReader();
                        dataTable.Load(reader);

                        // cmd.ExecuteNonQuery();
                        //connection.Close();
                        //return true;
                    }
                }
                return dataTable;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Console.WriteLine("ha2 " + ex);
                return null;
            }
        }
    }


}
