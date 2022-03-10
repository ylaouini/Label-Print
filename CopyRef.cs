using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Label_Print
{
    public partial class CopyRef : Form
    {
        public string epn;
        public CopyRef()
        {
            InitializeComponent();
        }

        private void btnDupplicate_Click(object sender, EventArgs e)
        {
            bool dupplicated = false;

            if (txtNewEpn.Text == "" || txtNewEpn.Text == null || txtNewEpn.Text == "EPN")
            {
                MessageBox.Show("EPN cannot be empty ", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string query = "INSERT INTO [dbo].[REFERENCE]  ([EPN] ,[CPN1] ,[CPN2] ,[CPN3]  ,[ALERT1] ,[ALERT2] ,[ALERT3] ,[RELEASE] ,[CUSTOMER] ,[LOT] ,[LEVEL_] ,[INDICE] ,[ETOILE] ,[PRFX] ,[ID_FAMILY] ,[OPERATOR] ,[OLL], [ID_ETIQUETTE]) " +
                "SELECT '"+txtNewEpn.Text+"' ,[CPN1] ,[CPN2] ,[CPN3]  ,[ALERT1] ,[ALERT2] ,[ALERT3] ,[RELEASE] ,[CUSTOMER] ,[LOT] ,[LEVEL_] ,[INDICE] ,[ETOILE] ,[PRFX] ,[ID_FAMILY] ,[OPERATOR] ,[OLL], [ID_ETIQUETTE] FROM [dbo].[REFERENCE] WHERE [EPN]='"+epn+"'";

            dupplicated = DBAccess.ExecuteQuery(query);

            if (dupplicated)
            {
                this.Close();
            }
        }

        private void CopyRef_Load(object sender, EventArgs e)
        {
            label1.Text = epn;
        }
    }
}
