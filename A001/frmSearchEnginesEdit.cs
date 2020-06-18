using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Net;

namespace LeadsExtractor
{
    public partial class frmSearchEnginesEdit : Form
    {
        private Form1 m_parent;
        private String m_strMode;
        private Int32 m_nItem;
        private String m_strSearchEngineName;
        private String m_strSearchEngineURL;
        private bool m_bIsEnabled;

        public frmSearchEnginesEdit(Form1 frm1, 
                                    String strMode, 
                                    Int32 nItem,
                                    String strSearchEngineName,
                                    String strSearchEngineURL,
                                    bool bIsEnabled)
        {
            InitializeComponent();

            m_parent = frm1;
            m_strMode = strMode;
            m_nItem = nItem;
            m_strSearchEngineName = strSearchEngineName;
            m_strSearchEngineURL = strSearchEngineURL;
            m_bIsEnabled = bIsEnabled;

            if (m_strMode == "REMOVE")
            {
                this.Text = "Remove Search Engine";
                txtSearchEngineName.Text = m_strSearchEngineName;
                txtSearchEngineName.ReadOnly = true;
                txtSearchEngineURL.Text = m_strSearchEngineURL;
                txtSearchEngineURL.ReadOnly = true;
                checkBox1.Checked = m_bIsEnabled;
                checkBox1.AutoCheck = false;
            }

            if (m_strMode == "UPDATE")
            {
                this.Text = "Update Search Engine";
                txtSearchEngineName.Text = m_strSearchEngineName;                
                txtSearchEngineURL.Text = m_strSearchEngineURL;                
                checkBox1.Checked = m_bIsEnabled;                  
            }

            if (m_strMode == "ADD")
            {
                this.Text = "Add Search Engine";
                txtSearchEngineName.Text = "";
                txtSearchEngineURL.Text = "";
                checkBox1.Checked = false;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            if ((txtSearchEngineName.TextLength < 4) && ((m_strMode == "UPDATE")||(m_strMode == "ADD")) )
            {


                MessageBox.Show("Search Engine Name should be at least 4 chars!",
                                "Input Error #1!",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Exclamation);
            }
            else if (((!CheckUrlStatus(txtSearchEngineURL.Text)) || (txtSearchEngineURL.TextLength < 5)) && ((m_strMode == "UPDATE")||(m_strMode == "ADD")))
            {


                MessageBox.Show("Search Engine URL is invalid!",
                                "Input Error #2!",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Exclamation);

            }
            else if ((m_parent.CheckIfSearchEngineExistsInSearchEngineListView(txtSearchEngineName.Text)>0) && (m_strMode == "ADD"))
            {

                MessageBox.Show("Such Search Engine Name already exists!",
                                "Input Error #3!",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Exclamation);
            }
            else if ((m_parent.CheckIfSearchEngineURLExistsInSearchEngineListView(txtSearchEngineURL.Text)>0) && (m_strMode == "ADD"))
            {

                MessageBox.Show("Such Search Engine URL already exists!",
                                "Input Error #4!",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Exclamation);
            }
            else if ((m_parent.CheckIfSearchEngineExistsInSearchEngineListView(txtSearchEngineName.Text) >1) && (m_strMode == "UPDATE"))
            {

                MessageBox.Show("Such Search Engine Name already exists!",
                                "Input Error #5!",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Exclamation);
            }
            else if ((m_parent.CheckIfSearchEngineURLExistsInSearchEngineListView(txtSearchEngineURL.Text) >1) && (m_strMode == "UPDATE"))
            {

                MessageBox.Show("Such Search Engine URL already exists!",
                                "Input Error #6!",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Exclamation);
            }
            else
            {

                // ok - remove item
                if (m_strMode == "REMOVE")
                {
                    m_parent.RemoveItem_listOfEngines(m_nItem);
                }

                // ok - update item
                if (m_strMode == "UPDATE")
                {
                    m_parent.UpdateItem_listOfEngines(m_nItem,
                                                      txtSearchEngineName.Text,
                                                      txtSearchEngineURL.Text,
                                                      checkBox1.Checked);
                }

                // ok - add item
                if (m_strMode == "ADD")
                {
                    m_parent.AddItem_listOfEngines(m_nItem,
                                                      txtSearchEngineName.Text,
                                                      txtSearchEngineURL.Text,
                                                      checkBox1.Checked);
                }

                this.Close();

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // cancel
            this.Close();
        }

        private void frmSearchEnginesEdit_Load(object sender, EventArgs e)
        {

        }


        //
        // CheckUrlStatus
        //

        protected bool CheckUrlStatus(string Website)
        {
            try
            {
                var request = WebRequest.Create(Website) as HttpWebRequest;
                request.Method = "HEAD";
                using (var response = (HttpWebResponse)request.GetResponse())
                {
                    return response.StatusCode == HttpStatusCode.OK;
                }
            }
            catch
            {
                return false;
            }
        }

        private void txtSearchEngineURL_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                button1.PerformClick();
            }
        }

        private void txtSearchEngineName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                button1.PerformClick();
            }
        }

 

    }
}
