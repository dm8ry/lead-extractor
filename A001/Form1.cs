using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;
using System.Net;
using System.Threading;
using System.Globalization;
using System.Xml;
using System.Reflection;
using System.Diagnostics;


namespace LeadsExtractor
{

    public partial class Form1 : Form
    {

        static Int32 nState = -1; // -1 = Stop; 0 = Pause; 1 = Start

        static List<SearchEngine> listOfEngines = new List<SearchEngine>();
        static List<Lead> listOfLeads = new List<Lead>();
        static List<URLprocessed> listOfURLsProcessed = new List<URLprocessed>();
        static List<URLprocessed> listOfURLsTOBEProcessed = new List<URLprocessed>();
        
        static String strVersion = "2.00";
        static String strLicenseToName = "Licensed To: Jimmy Braun";
        static String strNavigateLink = "http://www.microsoft.com";

        static Int32 nLeadId = 0;
        
        static String strProcessState = "";
        static Int32 nProcessStateTheSame = 0;        

        static String strAutosaveFolder = Directory.GetCurrentDirectory().ToString() + "\\" + "autosave_leads";
        static bool IsRefreshLeadsList = true;

        static Int32 nMinutesOfRun = 0;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
              this.SetStyle(ControlStyles.DoubleBuffer, true);
              this.SetStyle(ControlStyles.AllPaintingInWmPaint, true);
              this.SetStyle(ControlStyles.UserPaint, true);
              this.SetStyle(ControlStyles.ResizeRedraw, true);
              /*   
               var method1 = typeof(Control).GetMethod("SetStyle", BindingFlags.Instance | BindingFlags.NonPublic);
               method1.Invoke(listView2, new object[] { ControlStyles.OptimizedDoubleBuffer, true});

               var method2 = typeof(Control).GetMethod("SetStyle", BindingFlags.Instance | BindingFlags.NonPublic);
               method2.Invoke(listView3, new object[] { ControlStyles.OptimizedDoubleBuffer, true });
               */
          
            load_xml_settings_on_start();

            populateSearchEngines();
          
            // update buttons state - Stop Pause Start
            ButtonsState();
                          
        }

        /////////////////////////////////////////////////////////////////
        //
        // populate search engines
        //

        void populateSearchEngines()
        {
            listOfEngines.Clear();

            for (Int32 idx = 0; idx < listViewSearchEngines.Items.Count; idx++)
            {                                

               listOfEngines.Add(new SearchEngine(listViewSearchEngines.Items[idx].SubItems[0].Text, 
                                                  listViewSearchEngines.Items[idx].SubItems[1].Text, 
                                                  listViewSearchEngines.Items[idx].Checked));
            }

            listOfEngines = listOfEngines.OrderBy(o => o.getEngineName()).ToList();
            refreshSearchEnginesListView();
        }

        ///////////////////////////////////////////////////////////////
        //
        // SplitWords
        //

        static string[] SplitWords(string s)
        {
            return Regex.Split(s, @"\W+");
        }


        //////////////////////////////////////////////////////////////
        //
        // Start button pressed
        //

        private void button2_Click(object sender, EventArgs e)
        {
            start_processing(); 
        }

        ///////////////////////////////////////////////////////////////
        //
        //

        private void start_processing()
        {

            string[] strKeyWords = SplitWords(txtKeywords.Text);
            string strToSearch = "";

            foreach (string s in strKeyWords)
            {
                strToSearch = strToSearch + s + '+';
            }

            strToSearch = strToSearch.Substring(0, strToSearch.Length - 1);            

            populateSearchEngines();

            // initial 
            for (int idx = 0; idx < listOfEngines.Count; idx++)
            {
                if (listOfEngines[idx].getIsEnabled() == true)
                {
                    listOfURLsTOBEProcessed.Add(new URLprocessed(listOfEngines[idx].getEngineURL() + strToSearch,
                                                                    0,
                                                                    0,
                                                                    0,
                                                                    0,
                                                                    listOfEngines[idx].getEngineURL(),
                                                                    ""));
                }
            }

            nState = 1;
            ButtonsState();  
        }

        //////////////////////////////////////////////////////////////////
        //
        // Function To Process URL To Extract Emails
        //
        //
         
        private void ProcessURL(object messageObj)
        {

            // process single URL
            // no care duplicates
            

            try {                

                if (nState != 1) return;                

                this.Invoke((MethodInvoker)delegate
                {
                    toolStripProgressBar1.Value = 0;
                    toolStripStatusLabelProgress.Text = "00" + toolStripProgressBar1.Value.ToString();
                });

                String strCurrentProcessState = "L1" + toolStripStatusLabel2.Text +
                                                "L2" + toolStripStatusLabel5.Text +
                                                "L3" + toolStripStatusLabel9.Text +
                                                "L4" + toolStripStatusLabel12.Text +
                                                "L5" + toolStripStatusLabel15.Text +
                                                "L6" + toolStripStatusLabel18.Text + "L7";

                if (strCurrentProcessState == strProcessState)
                {
                    nProcessStateTheSame++;
                }
                else
                {
                    strProcessState = strCurrentProcessState;
                    nProcessStateTheSame = 0;
                }

                Int32 nNothingProcessAttempts = 0;
                    
                this.Invoke((MethodInvoker)delegate
                {
                    nNothingProcessAttempts = Convert.ToInt32(cbxNothingToProcessAtmpts.Text);
                });

                if ((nProcessStateTheSame > nNothingProcessAttempts) && (nState == 1))
                {
                    nState = -1;
                    ButtonsState();

                    MessageBox.Show("It looks nothing to process." + Environment.NewLine
                        + "The process has stopped." + Environment.NewLine
                        + "Try to change criteria(s).", "LeadsExtractor");
                }
                    
                URLprocessed objURLToProcess = (URLprocessed)messageObj;

                String strInpURL = objURLToProcess.getURLProcessed();
                Int32 nCurrentDepth = objURLToProcess.getnDepthLevel();
                String strCurrentParent = objURLToProcess.getSearchParent();
                String strCurrentHistoryChain = objURLToProcess.getSearchHistoryChain();
                String strChangeHistoryChain = strCurrentHistoryChain + "-->" + strInpURL;

                Int32 nMaxSearchDepth = 0;

                this.Invoke((MethodInvoker)delegate
                {
                   nMaxSearchDepth = Convert.ToInt32(comboMaxSearchDepth.Text);                    
                });

                if (nCurrentDepth > nMaxSearchDepth)
                    return;

                WebClient client = new WebClient();
                String htmlCode = client.DownloadString(strInpURL);

                bool bCheckBox5Status = false;
                string strTextBox2 = "";

                this.Invoke((MethodInvoker)delegate
                {
                    strTextBox2 = textBox2.Text.ToString().ToUpper();
                    bCheckBox5Status = checkBox5.Checked;
                });

                if ((htmlCode.ToUpper().IndexOf(strTextBox2) == -1) && 
                    (bCheckBox5Status == true) && 
                    (strTextBox2.Length > 0))
                {
                    return;
                }              
             
                Int64 nSizeOfFileToProcess = System.Text.ASCIIEncoding.Unicode.GetByteCount(htmlCode);
                Int64 nMaxWebPageSize = 0;

                this.Invoke((MethodInvoker)delegate
                {
                    nMaxWebPageSize = Convert.ToInt64(cbxMaxWebPagesSize.Text) * 1024;                
                });

                if (nSizeOfFileToProcess > nMaxWebPageSize)
                    return;

                // progress bar 5
                this.Invoke((MethodInvoker)delegate
                {
                toolStripProgressBar1.Value = 5;
                toolStripStatusLabelProgress.Text = "00" + toolStripProgressBar1.Value.ToString();
                });

                Regex regEx = new Regex(@"[\n|\r]+");
                htmlCode = regEx.Replace(htmlCode, Environment.NewLine);
                htmlCode = htmlCode.Replace("&#45;", "-");
                htmlCode = htmlCode.Replace("&#46;", ".");
                htmlCode = htmlCode.Replace("&#48;", "0");
                htmlCode = htmlCode.Replace("&#49;", "1");
                htmlCode = htmlCode.Replace("&#50;", "2");
                htmlCode = htmlCode.Replace("&#51;", "3");
                htmlCode = htmlCode.Replace("&#52;", "4");
                htmlCode = htmlCode.Replace("&#53;", "5");
                htmlCode = htmlCode.Replace("&#54;", "6");
                htmlCode = htmlCode.Replace("&#55;", "7");
                htmlCode = htmlCode.Replace("&#56;", "8");
                htmlCode = htmlCode.Replace("&#57;", "9");
                htmlCode = htmlCode.Replace("&#64;", "@");
                htmlCode = htmlCode.Replace("&#65;", "A");
                htmlCode = htmlCode.Replace("&#66;", "B");
                htmlCode = htmlCode.Replace("&#67;", "C");
                htmlCode = htmlCode.Replace("&#68;", "D");
                htmlCode = htmlCode.Replace("&#69;", "E");
                htmlCode = htmlCode.Replace("&#70;", "F");
                htmlCode = htmlCode.Replace("&#71;", "G");
                htmlCode = htmlCode.Replace("&#72;", "H");
                htmlCode = htmlCode.Replace("&#73;", "I");
                htmlCode = htmlCode.Replace("&#74;", "J");
                htmlCode = htmlCode.Replace("&#75;", "K");
                htmlCode = htmlCode.Replace("&#76;", "L");
                htmlCode = htmlCode.Replace("&#77;", "M");
                htmlCode = htmlCode.Replace("&#78;", "N");
                htmlCode = htmlCode.Replace("&#79;", "O");
                htmlCode = htmlCode.Replace("&#80;", "P");
                htmlCode = htmlCode.Replace("&#81;", "Q");
                htmlCode = htmlCode.Replace("&#82;", "R");
                htmlCode = htmlCode.Replace("&#83;", "S");
                htmlCode = htmlCode.Replace("&#84;", "T");
                htmlCode = htmlCode.Replace("&#85;", "U");
                htmlCode = htmlCode.Replace("&#86;", "V");
                htmlCode = htmlCode.Replace("&#87;", "W");
                htmlCode = htmlCode.Replace("&#88;", "X");
                htmlCode = htmlCode.Replace("&#89;", "Y");
                htmlCode = htmlCode.Replace("&#90;", "Z");
                htmlCode = htmlCode.Replace("&#95;", "_");
                htmlCode = htmlCode.Replace("&#97;", "a");
                htmlCode = htmlCode.Replace("&#98;", "b");
                htmlCode = htmlCode.Replace("&#99;", "c");
                htmlCode = htmlCode.Replace("&#100;", "d");
                htmlCode = htmlCode.Replace("&#101;", "e");
                htmlCode = htmlCode.Replace("&#102;", "f");
                htmlCode = htmlCode.Replace("&#103;", "g");
                htmlCode = htmlCode.Replace("&#104;", "h");
                htmlCode = htmlCode.Replace("&#105;", "i");
                htmlCode = htmlCode.Replace("&#106;", "j");
                htmlCode = htmlCode.Replace("&#107;", "k");
                htmlCode = htmlCode.Replace("&#108;", "l");
                htmlCode = htmlCode.Replace("&#109;", "m");
                htmlCode = htmlCode.Replace("&#110;", "n");
                htmlCode = htmlCode.Replace("&#111;", "o");
                htmlCode = htmlCode.Replace("&#112;", "p");
                htmlCode = htmlCode.Replace("&#113;", "q");
                htmlCode = htmlCode.Replace("&#114;", "r");
                htmlCode = htmlCode.Replace("&#115;", "s");
                htmlCode = htmlCode.Replace("&#116;", "t");
                htmlCode = htmlCode.Replace("&#117;", "u");
                htmlCode = htmlCode.Replace("&#118;", "v");
                htmlCode = htmlCode.Replace("&#119;", "w");
                htmlCode = htmlCode.Replace("&#120;", "x");
                htmlCode = htmlCode.Replace("&#121;", "y");
                htmlCode = htmlCode.Replace("&#122;", "z");

                // progress bar 10
               this.Invoke((MethodInvoker)delegate
                {
                    toolStripProgressBar1.Value = 10;
                    toolStripStatusLabelProgress.Text = "0" + toolStripProgressBar1.Value.ToString();
                });

                // extract email leads
                Regex emailRegex = new Regex(txtEmailRegEx.Text,
                            RegexOptions.IgnoreCase);

                MatchCollection emailMatches = emailRegex.Matches(htmlCode);

                Int32 nEmailsCount = 0;

                    this.Invoke((MethodInvoker)delegate
                    {

                    foreach (Match emailMatch in emailMatches)
                    {
                        bool bFlag = false;

                        for (Int32 idx = 0; idx < lv_ExcludeEmailsWithName.Items.Count; idx++)
                        {
                            if (emailMatch.Value.ToUpper().IndexOf(lv_ExcludeEmailsWithName.Items[idx].Text.ToUpper()) != -1) 
                                bFlag = true;
                        }

                        for (Int32 idx = 0; idx < lv_ExcludeEmailsWithDomain.Items.Count; idx++)
                        {
                            if ((emailMatch.Value.ToUpper().IndexOf(lv_ExcludeEmailsWithDomain.Items[idx].Text.ToUpper()) != -1)
                                    &&
                                (emailMatch.Value.ToUpper().IndexOf(lv_ExcludeEmailsWithDomain.Items[idx].Text.ToUpper()) > emailMatch.Value.ToUpper().IndexOf("@")))
                                bFlag = true;
                        }

                        if (
                            (emailMatch.Value.Length <= Convert.ToInt32(cbxMaxEmailSize.Text))
                            &&
                            bFlag == false
                            )

                        {
                            listOfLeads.Add(new Lead(++nLeadId,
                                                   LeadType.Email,
                                                   emailMatch.Value,
                                                   strInpURL));

                            nEmailsCount++;
                        }
                    }
                });

                // progress bar 20
                this.Invoke((MethodInvoker)delegate
                {
                   toolStripProgressBar1.Value = 20;
                   toolStripStatusLabelProgress.Text = "0" + toolStripProgressBar1.Value.ToString();
                });


                // extract phone leads
                Regex phoneRegex = new Regex(txtPhoneRegEx.Text,
                            RegexOptions.IgnoreCase);

                MatchCollection phoneMatches = phoneRegex.Matches(htmlCode);

                Int32 nPhonesCount = 0;

                this.Invoke((MethodInvoker)delegate
                {
                foreach (Match phoneMatch in phoneMatches)
                {
                    if (phoneMatch.Value.Length <= Convert.ToInt32(cbxMaxPhoneSize.Text))
                    {
                        listOfLeads.Add(new Lead(++nLeadId,
                                                LeadType.Phone,
                                                phoneMatch.Value,
                                                strInpURL));

                        nPhonesCount++;
                    }
                }
                });

                // progress bar 40
                this.Invoke((MethodInvoker)delegate
                {
                  toolStripProgressBar1.Value = 40;
                  toolStripStatusLabelProgress.Text = "0" + toolStripProgressBar1.Value.ToString();
                });


                // extract skype leads 
                Regex skypeRegex = new Regex(txtSkypeRegEx.Text,
                            RegexOptions.IgnoreCase);

                MatchCollection skypeMatches = skypeRegex.Matches(htmlCode);

                Int32 nSkypesCount = 0;

                this.Invoke((MethodInvoker)delegate
                {

                foreach (Match skypeMatch in skypeMatches)
                {
                    if (skypeMatch.Value.Length <= Convert.ToInt32(cbxMaxSkypeSize.Text))
                    {
                        listOfLeads.Add(new Lead(++nLeadId,
                                                LeadType.Skype,
                                                skypeMatch.Value,
                                                strInpURL));

                        nSkypesCount++;
                    }
                }

                });

                listOfURLsProcessed.Add(new URLprocessed(strInpURL,
                                                        nEmailsCount,
                                                        nPhonesCount,
                                                        nSkypesCount,
                                                        nCurrentDepth,
                                                        strCurrentParent,
                                                        strChangeHistoryChain));
                

                // progress bar 50
                this.Invoke((MethodInvoker)delegate
                {
                toolStripProgressBar1.Value = 50;
                toolStripStatusLabelProgress.Text = "0" + toolStripProgressBar1.Value.ToString();               
                });

                // extract urls
                Regex regEx2 = new Regex(@"[\n|\r]+");
                htmlCode = regEx2.Replace(htmlCode, Environment.NewLine);

                Regex urlsRegex = new Regex(@"(http|ftp|https)://([\w+?\.\w+])+([a-zA-Z0-9\~\!\@\#\$\%\^\&\*\(\)_\-\=\+\\\/\?\.\:\;\'\,]*)?",
                            RegexOptions.IgnoreCase);

                MatchCollection urlsMatches = urlsRegex.Matches(htmlCode);
               
                // progress bar 80
                this.Invoke((MethodInvoker)delegate
                {
                toolStripProgressBar1.Value = 80;
                toolStripStatusLabelProgress.Text = "0" + toolStripProgressBar1.Value.ToString();
                });

                this.Invoke((MethodInvoker)delegate
                {
                     foreach (Match urlsMatch in urlsMatches)
                     {
                         
                         bool bFlag = false;

                         for (Int32 idx = 0; idx < lv_ExcludeURLsWith.Items.Count; idx++)
                         {                            
                             if (urlsMatch.Value.ToUpper().IndexOf(lv_ExcludeURLsWith.Items[idx].Text.ToUpper()) != -1)
                                 bFlag = true;
                         }

                         if (bFlag == false)
                         {
                            Int32 nMaxBuffSzPreparedURL = 0;
                             nMaxBuffSzPreparedURL = Convert.ToInt32(comboMaxBufferSzPrepareURL.Text);
                             if (listOfURLsTOBEProcessed.Count < nMaxBuffSzPreparedURL) // max buffers size for listOfURLsTOBEProcessed
                             {
                                   if (
                                     (checkBox4.Checked == true)
                                     &&
                                     (urlsMatch.Value.ToString().ToUpper().IndexOf(txtKWdomainName.Text.ToUpper()) != -1)
                                     &&
                                     (txtKWdomainName.Text.Length >0)
                                     ||
                                     (checkBox4.Checked == false)
                                     )
                                 {

                                     if ((nCurrentDepth + 1) <= Convert.ToInt32(comboMaxSearchDepth.Text))
                                     { 

                                         listOfURLsTOBEProcessed.Add(new URLprocessed(urlsMatch.Value,
                                                                                                     0,
                                                                                                     0,
                                                                                                     0,
                                                                                                     nCurrentDepth + 1,
                                                                                                     strCurrentParent,
                                                                                                     strChangeHistoryChain));
                                          }
                                     } 

                             } // if (listOfURLsTOBEProcessed.Count < nMaxBuffSzPreparedURL)

                         } // if (bFlag == false)
                     } // foreach

                     listOfURLsTOBEProcessed = listOfURLsTOBEProcessed.Distinct().ToList();
                });

                this.Invoke((MethodInvoker)delegate
                {
                    toolStripProgressBar1.Value = 100;
                    toolStripStatusLabelProgress.Text = toolStripProgressBar1.Value.ToString();
                            
                });

                }
                catch (Exception ex)
                {
                   // MessageBox.Show(ex.ToString());
                }            
              
        }

        //
        // Update Start Pause Stop buttons state
        private void ButtonsState()
        { 
            switch (nState)
            {
                case -1: // stop
                    button2.Enabled = true;
                    button3.Enabled = false;
                    button4.Enabled = false;
                    timer1.Enabled = false;
                    timer2.Enabled = false;
                    timer3.Enabled = false;
                    timer3.Interval = 60 * 1000 * Convert.ToInt32(cbxAutoSaveResults.Text);
                    timer4.Enabled = true;
                    startToolStripMenuItem.Enabled = true;
                    pauseToolStripMenuItem.Enabled = false;
                    stopToolStripMenuItem.Enabled = false;
                    nLeadId = 0;
                    lblState.Text = "Stop";
                    break;

                case 0: // pause
                    button2.Enabled = false;
                    startToolStripMenuItem.Enabled = false;
                    button3.Enabled = true;
                    pauseToolStripMenuItem.Enabled = true;
                    if (button3.Text == "Resume")
                    { 
                        button3.Text = "Pause";
                        pauseToolStripMenuItem.Text = "Pause";
                        button4.Enabled = true;
                        stopToolStripMenuItem.Enabled = true;
                        timer1.Enabled = true;
                        timer2.Enabled = true;
                        timer3.Enabled = true;
                        timer3.Interval = 60 * 1000 * Convert.ToInt32(cbxAutoSaveResults.Text);
                        timer4.Enabled = true;
                        nState = 1;
                        lblState.Text = "Run";
                    }
                    else
                    {
                        button3.Text = "Resume";
                        pauseToolStripMenuItem.Text = "Resume";
                        button4.Enabled = false;
                        stopToolStripMenuItem.Enabled = false;
                        timer1.Enabled = false;
                        timer2.Enabled = false;
                        timer3.Enabled = false;
                        timer3.Interval = 60 * 1000 * Convert.ToInt32(cbxAutoSaveResults.Text);
                        timer4.Enabled = false;
                        nState = 0;
                        lblState.Text = "Pause";
                    }
                    
                    break;

                case 1: // start
                    button2.Enabled = false;
                    startToolStripMenuItem.Enabled = false;
                    button3.Enabled = true;
                    button4.Enabled = true;
                    stopToolStripMenuItem.Enabled = true;
                    timer1.Enabled = true;
                    timer2.Enabled = true;
                    timer3.Enabled = true;
                    timer3.Interval = 60 * 1000 * Convert.ToInt32(cbxAutoSaveResults.Text);
                    timer4.Enabled = true;
                    pauseToolStripMenuItem.Enabled = true;
                    lblState.Text = "Run";
                    break;

                default:
                    button2.Enabled = false;
                    button3.Enabled = false;
                    button4.Enabled = false;
                    stopToolStripMenuItem.Enabled = false;
                    startToolStripMenuItem.Enabled = false;
                    pauseToolStripMenuItem.Enabled = false;
                    timer1.Enabled = false;
                    timer2.Enabled = false;
                    timer3.Enabled = false;
                    timer3.Interval = 60 * 1000 * Convert.ToInt32(cbxAutoSaveResults.Text);
                    timer4.Enabled = false;
                    lblState.Text = "Unknown";
                    break;
            }
              
        }

        private void button3_Click(object sender, EventArgs e)
        {         
            nState = 0;
            ButtonsState();          
        }

        private void button4_Click(object sender, EventArgs e)
        {           
            nState = -1;      
            ButtonsState(); 
        }

        // tick of timer to do processing

        private void timer1_Tick(object sender, EventArgs e)
        {

            if (nState == 1)
            {
                Int32 nParallelDegree = Convert.ToInt32(comboDegreeParallelism.Text);
                for (int idx = 0; idx < nParallelDegree; idx++)
                {
                    // degree parallelism 
                    if (listOfURLsTOBEProcessed.Count > 0)
                    {
                        URLprocessed objURLToProcess = listOfURLsTOBEProcessed[0];
                        listOfURLsTOBEProcessed.RemoveAt(0);
                        Thread tProcessURL = new Thread(ProcessURL);
                        tProcessURL.Start(objURLToProcess);
                     }
                }
               
            } // if

            if (listOfURLsTOBEProcessed.Count == 0)
                nProcessStateTheSame++;
            else
                nProcessStateTheSame = 0;

            Int32 nNothingProcessAttempts = 0;

            nNothingProcessAttempts = Convert.ToInt32(cbxNothingToProcessAtmpts.Text);

            if ((nProcessStateTheSame > nNothingProcessAttempts) && (nState == 1))
            {
                nState = -1;
                ButtonsState();

                MessageBox.Show("It looks nothing to process." + Environment.NewLine
                    + "The process has stopped." + Environment.NewLine
                    + "Try to change criteria(s).", "LeadsExtractor");
            }
              
        }

        //////////////////////////////////////////////////
        //
        // copy all the leads to clipboard
        //

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Clipboard.Clear();
            StringBuilder buffer = new StringBuilder();

            // Setup the columns

            for (int i = 0; i < this.listView3.Columns.Count; i++)
            {
                buffer.Append(this.listView3.Columns[i].Text);
                buffer.Append("\t");
            }
            buffer.Append(Environment.NewLine);

            // Build the data row by row

            for (int i = 0; i < this.listView3.Items.Count; i++)
            {

                for (int j = 0; j < this.listView3.Columns.Count; j++)
                {
                    buffer.Append(this.listView3.Items[i].SubItems[j].Text);
                    buffer.Append("\t");
                }
                buffer.Append(Environment.NewLine);
            }

            Clipboard.SetText(buffer.ToString());
                
        }

        // export all to excel
        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {               
           ExportToExcel(chkboxAddHeaderExportExcel.Checked, false);            
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void leadsExtractedToExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {

            Int32 nOfItems = listView3.Items.Count;

            if (nOfItems == 0)
            {
                MessageBox.Show("No leads to export!", "LeadsExtractor");
            }
            else
            {
                ExportToExcel(chkboxAddHeaderExportExcel.Checked, false);
            }
        }


        /////////////////////////////////////////////////////////////////
        //
        // export listview of leads to excel file
        //
        // nWithHeader: true - with header; false - without header
        // bJustSelected: true - just selected leads; false - all the leads
        //

        private void ExportToExcel(bool nWithHeader, bool bJustSelected)
        {                       
            try
            {
                saveFileDialog1.RestoreDirectory = true;
                saveFileDialog1.Title = "Export To Excel File";
                saveFileDialog1.DefaultExt = "xlsx";
                saveFileDialog1.Filter = "Excel files (*.xlsx) |*.xlsx|All files (*.*)|*.*";
                saveFileDialog1.FilterIndex = 1;

                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {

                    Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                    ExcelApp.Application.Workbooks.Add(Type.Missing);                    
                    ExcelApp.Columns.ColumnWidth = 30;


                    if (bJustSelected == true) // export just selected leads to excel
                    {
                        Int32 nOfSelectedItems = listView3.SelectedItems.Count;

                        if (nWithHeader == true)
                        {

                            ExcelApp.Cells[1, 1] = "Lead Id";
                            ExcelApp.Cells[1, 2] = "Lead Create Date";
                            ExcelApp.Cells[1, 3] = "Lead Type";
                            ExcelApp.Cells[1, 4] = "Lead Content";
                            ExcelApp.Cells[1, 5] = "Source URL of the Lead";

                            for (Int32 s = 0; s < nOfSelectedItems; s++)
                            {
                                ListViewItem item = listView3.SelectedItems[s];
 
                                for (int j = 0; j < listView3.Columns.Count; j++)
                                {
                                    ExcelApp.Cells[s + 1 + 1, j + 1] = item.SubItems[j].Text.ToString();
                                }
                            }

                        }
                        else
                        {
                            for (Int32 s = 0; s < nOfSelectedItems; s++)
                            {
                                ListViewItem item = listView3.SelectedItems[s];

                                for (int j = 0; j < listView3.Columns.Count; j++)
                                {
                                    ExcelApp.Cells[s + 1, j + 1] = item.SubItems[j].Text.ToString();
                                }
                            }
                        }
                    }
                    else // export all the leads to excel
                    {

                        if (nWithHeader == true)
                        {

                            ExcelApp.Cells[1, 1] = "Lead Id";
                            ExcelApp.Cells[1, 2] = "Lead Create Date";
                            ExcelApp.Cells[1, 3] = "Lead Type";
                            ExcelApp.Cells[1, 4] = "Lead Content";
                            ExcelApp.Cells[1, 5] = "Source URL of the Lead";

                            for (int i = 0; i < listView3.Items.Count; i++)
                            {
                                for (int j = 0; j < listView3.Columns.Count; j++)
                                {
                                    ExcelApp.Cells[i + 1 + 1, j + 1] = listView3.Items[i].SubItems[j].Text.ToString();
                                }
                            }

                        }
                        else
                        {
                            for (int i = 0; i < listView3.Items.Count; i++)
                            {
                                for (int j = 0; j < listView3.Columns.Count; j++)
                                {
                                    ExcelApp.Cells[i + 1, j + 1] = listView3.Items[i].SubItems[j].Text.ToString();
                                }
                            }
                        }

                    }

                    ExcelApp.ActiveWorkbook.SaveAs(saveFileDialog1.FileName);
                    ExcelApp.Quit();
                }
            }
            catch (Exception ex)
            {
                this.Close();
            } 
        }

        private void autoSaveLeads(String strFileName)
        {
            try
            {
                if (listView3.Items.Count == 0) return;

                    Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                    ExcelApp.Application.Workbooks.Add(Type.Missing);
                    ExcelApp.Columns.ColumnWidth = 30;
                   
                    ExcelApp.Cells[1, 1] = "Lead Id";
                    ExcelApp.Cells[1, 2] = "Lead Create Date";
                    ExcelApp.Cells[1, 3] = "Lead Type";
                    ExcelApp.Cells[1, 4] = "Lead Content";
                    ExcelApp.Cells[1, 5] = "Source URL of the Lead";

                    for (int i = 0; i < listView3.Items.Count; i++)
                    {
                        for (int j = 0; j < listView3.Columns.Count; j++)
                        {
                            ExcelApp.Cells[i + 1 + 1, j + 1] = listView3.Items[i].SubItems[j].Text.ToString();
                        }
                    }

                    ExcelApp.ActiveWorkbook.SaveAs(strFileName);
                    ExcelApp.Quit();                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            } 
        }

        private void startToolStripMenuItem_Click(object sender, EventArgs e)
        {
            start_processing(); 
        }

        private void pauseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            nState = 0;
            ButtonsState();  
        }

        private void stopToolStripMenuItem_Click(object sender, EventArgs e)
        {        
            nState = -1;
            ButtonsState();              
        }

        private void restoreDefaultsToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void removeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //
            // remove SearchEngine 

            // get selected item number
            if (listViewSearchEngines.SelectedItems.Count == 1)
            {
                Int32 nSelectedItem = listViewSearchEngines.SelectedItems[0].Index;

                frmSearchEnginesEdit wDialog = new frmSearchEnginesEdit(this,
                                                                        "REMOVE",
                                                                        nSelectedItem,
                                                                        listOfEngines[nSelectedItem].getEngineName(),
                                                                        listOfEngines[nSelectedItem].getEngineURL(),
                                                                        listOfEngines[nSelectedItem].getIsEnabled());
                wDialog.ShowDialog();                
            }
            refreshSearchEnginesListView();
        }


        private void changeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //
            // update SearchEngine details

            // get selected item number
            if (listViewSearchEngines.SelectedItems.Count == 1)
            {
                Int32 nSelectedItem = listViewSearchEngines.SelectedItems[0].Index;

                frmSearchEnginesEdit wDialog = new frmSearchEnginesEdit(this,
                                                                        "UPDATE",
                                                                        nSelectedItem,
                                                                        listOfEngines[nSelectedItem].getEngineName(),
                                                                        listOfEngines[nSelectedItem].getEngineURL(),
                                                                        listOfEngines[nSelectedItem].getIsEnabled());
                wDialog.ShowDialog();
            }

            refreshSearchEnginesListView();
        }

        ////////////////////////////////////////////////////////
        //
        // refreshSearchEnginesListView()
        //
        //

        private void refreshSearchEnginesListView()
        {
            listViewSearchEngines.Items.Clear();

            for (int idx = 0; idx < listOfEngines.Count; idx++)
            {
                ListViewItem oItem = new ListViewItem();

                if (listOfEngines[idx].getIsEnabled() == true)

                    oItem.Checked = true;
                else
                    oItem.Checked = false;

                oItem.Text = listOfEngines[idx].getEngineName();
                oItem.SubItems.Add(listOfEngines[idx].getEngineURL());
                listViewSearchEngines.Items.Add(oItem);
            }
        }

        //
        // aboutToolStripMenuItem_Click
        //

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //  show about window
            AboutBox1 wndAbout = new AboutBox1(strVersion);
            wndAbout.ShowDialog();   
        }

        //
        // RemoveItem_listOfEngines
        //

        public void RemoveItem_listOfEngines(Int32 nItem)
        {
            listOfEngines.RemoveAt(nItem);
        }

        //
        // UpdateItem_listOfEngines
        //

        public void UpdateItem_listOfEngines(Int32 nItem,
                                             String strSearchEngineName,
                                             String strSearchEngineURL,
                                             bool bIsEnabled)
        {
            listOfEngines[nItem].setEngineName(strSearchEngineName);
            listOfEngines[nItem].setEngineURL(strSearchEngineURL);
            listOfEngines[nItem].setIsEnabled(bIsEnabled);

            listOfEngines = listOfEngines.OrderBy(o => o.getEngineName()).ToList();
        }

        // Add new itme to SearchEngine ListView

        private void addToolStripMenuItem_Click(object sender, EventArgs e)
        {

            //
            // add a new SearchEngine item details
            
            frmSearchEnginesEdit wDialog = new frmSearchEnginesEdit(this,
                                                                    "ADD",
                                                                    -1,
                                                                    "",
                                                                    "",
                                                                    true);
            wDialog.ShowDialog();
           
            refreshSearchEnginesListView();

        }

        //////////////////////////////////////////////////////////////////
        //
        // AddItem_listOfEngines
        //

        public void AddItem_listOfEngines(Int32 nItem,
                                             String strSearchEngineName,
                                             String strSearchEngineURL,
                                             bool bIsEnabled)
        {            

            listOfEngines.Add(new SearchEngine(strSearchEngineName, 
                                            strSearchEngineURL, 
                                            bIsEnabled));

            listOfEngines = listOfEngines.OrderBy(o => o.getEngineName()).ToList();
        }

        /////////////////////////////////////////////////////////////////////
        //
        // restore defaults Search Engines
        //

        private void restoreDefaultsToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            // restore defaults

           // restoreSearchEngines();
        }


        ////////////////////////////////////////////////////////////////////////
        //
        // Check if SearchEngine name is already existing in the SearchEngine ListView
        //

        public Int32 CheckIfSearchEngineExistsInSearchEngineListView(String item)
        {
            Int32 bFound = 0;
            foreach (ListViewItem lvi in listViewSearchEngines.Items)
                if (String.Compare(item, lvi.Text, true, CultureInfo.CurrentCulture) == 0)
                {
                    bFound++;                    
                }
            return bFound;
        }

        //////////////////////////////////////////////////////////////////////////
        //
        // Check if SearchEngine name is already existing in the SearchEngine ListView
        //

        public Int32 CheckIfSearchEngineURLExistsInSearchEngineListView(String item)
        {
            Int32 bFound = 0;
            foreach (ListViewItem lvi in listViewSearchEngines.Items)
                if (String.Compare(item, lvi.SubItems[1].Text, true, CultureInfo.CurrentCulture) == 0)
                {
                    bFound++;
                    break;
                }
            return bFound;
        }


        private void readToolStripMenuItem_Click(object sender, EventArgs e)
        {
            /// read from xml
            /// 

        }


        private void statusStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        // clear collected leads and urls

        private void clearLeadsToolStripMenuItem_Click(object sender, EventArgs e)
        {

            DialogResult dialogResult = MessageBox.Show("Are You Sure?", "LeadsExtractor", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                
                listOfLeads.Clear();

                listView3.Items.Clear();

                listOfURLsTOBEProcessed.Clear();

                listOfURLsProcessed.Clear();

                listView2.Items.Clear();

                Int32 nNumberOfEmailsFound = 0;
                Int32 nNumberOfPhonesFound = 0;
                Int32 nNumberOfSkypesFound = 0;                    
                    
                toolStripStatusLabel2.Text = listOfURLsTOBEProcessed.Count().ToString();

                toolStripStatusLabel5.Text = listOfURLsProcessed.Count().ToString();

                toolStripStatusLabel9.Text = listOfLeads.Count().ToString();

                for (Int32 idx = 0; idx < listOfLeads.Count(); idx++)
                {
                    if (listOfLeads[idx].getLeadType() == LeadType.Email)
                        nNumberOfEmailsFound++;

                    if (listOfLeads[idx].getLeadType() == LeadType.Phone)
                        nNumberOfPhonesFound++;

                    if (listOfLeads[idx].getLeadType() == LeadType.Skype)
                        nNumberOfSkypesFound++;
                }

                toolStripStatusLabel12.Text = nNumberOfEmailsFound.ToString();

                toolStripStatusLabel15.Text = nNumberOfPhonesFound.ToString();

                toolStripStatusLabel18.Text = nNumberOfSkypesFound.ToString();
                
                chart1.Series["Parsed URLs"].Points.Clear();
                chart2.Series["Leads"].Points.Clear();
                chart3.Series["Emails"].Points.Clear();
                chart3.Series["Phones"].Points.Clear();
                chart3.Series["Skypes"].Points.Clear();

            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }

        }

        private void processSingleURLToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //
            //
            //
        }

        //////////////////////////////////////////////////////
        //
        // export just selected leads to excel
        //

        private void exportSelectedToExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // export selected leads to excel
            Int32 nOfSelectedItems = listView3.SelectedItems.Count;

            if (nOfSelectedItems == 0)
            {
                MessageBox.Show("No leads selected to export!", "LeadsExtractor");
            }
            else
            {
                ExportToExcel(chkboxAddHeaderExportExcel.Checked, true);
            }

        }

        private void exitToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            // Exit Application

            Application.Exit();
        }

        private void removeDuplicateLeads()
        {
            // remove duplicates Leads and URLs Processed

            for (int idx = 0; idx < listOfLeads.Count; idx++)
            {
                if (null == listOfLeads[idx])
                {
                    listOfLeads.RemoveAt(idx);
                }
            }

            List<Lead> lTemp;
            lTemp = listOfLeads.Distinct().ToList();
            listOfLeads = lTemp;

            //listOfLeads = listOfLeads.Distinct().ToList();
            listOfLeads = listOfLeads.GroupBy(item => item.getLeadContent().ToUpper()).Select(item => item.First()).ToList();

            for (int idx = 0; idx < listOfURLsTOBEProcessed.Count; idx++)
            {
                if (null == listOfURLsTOBEProcessed[idx])
                {
                    listOfURLsTOBEProcessed.RemoveAt(idx);
                }
            }
            listOfURLsTOBEProcessed = listOfURLsTOBEProcessed.Distinct().ToList();
            listOfURLsTOBEProcessed = listOfURLsTOBEProcessed.GroupBy(item => item.getURLProcessed().ToUpper()).Select(item => item.First()).ToList();

            for (int idx = 0; idx < listOfURLsProcessed.Count; idx++)
            {
                if (null == listOfURLsProcessed[idx])
                {
                    listOfURLsProcessed.RemoveAt(idx);
                }
            }
            listOfURLsProcessed = listOfURLsProcessed.Distinct().ToList();
            listOfURLsProcessed = listOfURLsProcessed.GroupBy(item => item.getURLProcessed().ToUpper()).Select(item => item.First()).ToList(); 

        }

        ///////////////////////////////////////////
        //
        //  save settings to xml
        //

        private void xMLToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            save_as_xml_settings();
        }

        ////////////////////////////////////////////////////////////
        //
        // restore default settings
        //

        private void restoreDefaultsToolStripMenuItem_Click_2(object sender, EventArgs e)
        {
            restoreDefaultSettings();
        }

        ////////////////////////////////////////////////////////////
        //
        // restore default settings
        //

        private void restoreDefaultSettingsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            restoreDefaultSettings();
        }

        ////////////////////////////////////////////////////////////
        //
        // restoreDefaultSettings
        //

        void restoreDefaultSettings()
        {

            lblLicenseDetails.Text = strLicenseToName;

            linkLabel1.TextAlign = ContentAlignment.MiddleCenter;
            linkLabel1.Location = new System.Drawing.Point(380, 411);
            linkLabel1.Text = "Leads Extractor" + " v" + strVersion + Environment.NewLine + "E-Solutions, 2016";            

            String strAppPath= Directory.GetCurrentDirectory().ToString();

            txtVersion.Text = strVersion;
            txtVersion.ReadOnly = true;

            txtAppPath.Text = strAppPath;
            txtAppPath.ReadOnly = true;

            txtSettingsXSDFileName.Text = "LeadsExtractorSettings.xsd";
            txtSettingsXSDFileName.ReadOnly = true;

            txtSettingsXSDFilePath.Text = strAppPath;
            txtSettingsXSDFilePath.ReadOnly = true;

            txtSettingsXMLDefaultFileName.Text = "LeadsExtractorSettings.xml";
            txtSettingsXMLDefaultFileName.ReadOnly = true;

            comboEngineFrequency.SelectedIndex = 2;
            comboDegreeParallelism.SelectedIndex = 4;
            comboMaxBufferSzPrepareURL.SelectedIndex = 8;

            txtDefaultPathForSettings.Text = txtAppPath.Text;
            txtDefaultPathForCampaigns.Text = txtAppPath.Text;

            txtDefaultPathForSettings.ReadOnly = true;
            txtDefaultPathForCampaigns.ReadOnly = true;

            comboMaxSearchDepth.SelectedIndex = 9;
            comboMaxLeadsToExtract.SelectedIndex = 4;
            cbxNothingToProcessAtmpts.SelectedIndex = 3;

            chkboxAddHeaderExportExcel.Checked = true;

            txtEmailRegEx.Text = @"[A-Z0-9._%+-]+@[A-Z0-9.-]{3,65}\.[A-Z]{2,4}";
            cbxMaxEmailSize.SelectedIndex = 4;

            txtPhoneRegEx.Text = @"[01]?[- .]?(\([2-9]\d{2}\)|[2-9]\d{2})[- .]?\d{3}[- .]?\d{4}";
            cbxMaxPhoneSize.SelectedIndex = 1;

            txtSkypeRegEx.Text = @"skype:\s*(?<name>[\w\.\s]*)";
            cbxMaxSkypeSize.SelectedIndex = 1;
            
            cbxMaxWebPagesSize.SelectedIndex = 5;
            cbxAutoSaveResults.SelectedIndex = 3;
            
            lv_ExcludeEmailsWithName.Items.Clear();
            lv_ExcludeEmailsWithName.Items.Add("admin");
            lv_ExcludeEmailsWithName.Items.Add("antispam");
            lv_ExcludeEmailsWithName.Items.Add("anti-spam");            
            lv_ExcludeEmailsWithName.Items.Add("no-contact");                        
            lv_ExcludeEmailsWithName.Items.Add("daemon");
            lv_ExcludeEmailsWithName.Items.Add("domain");
            lv_ExcludeEmailsWithName.Items.Add("ezine");
            lv_ExcludeEmailsWithName.Items.Add("abuse");
            lv_ExcludeEmailsWithName.Items.Add("info");
            lv_ExcludeEmailsWithName.Items.Add("sales");
            lv_ExcludeEmailsWithName.Items.Add("fidonet");
            lv_ExcludeEmailsWithName.Items.Add("fraud");
            lv_ExcludeEmailsWithName.Items.Add("-join");
            lv_ExcludeEmailsWithName.Items.Add("junk");
            lv_ExcludeEmailsWithName.Items.Add("-leave");
            lv_ExcludeEmailsWithName.Items.Add("listserv");
            lv_ExcludeEmailsWithName.Items.Add("mailer-daemon");
            lv_ExcludeEmailsWithName.Items.Add("majordomo");
            lv_ExcludeEmailsWithName.Items.Add("no-reply");
            lv_ExcludeEmailsWithName.Items.Add("no_reply");
            lv_ExcludeEmailsWithName.Items.Add("noreply");
            lv_ExcludeEmailsWithName.Items.Add("notice");
            lv_ExcludeEmailsWithName.Items.Add("nospam");
            lv_ExcludeEmailsWithName.Items.Add("privacy");
            lv_ExcludeEmailsWithName.Items.Add("rating");
            lv_ExcludeEmailsWithName.Items.Add("remove");
            lv_ExcludeEmailsWithName.Items.Add("removeme");
            lv_ExcludeEmailsWithName.Items.Add("resume");
            lv_ExcludeEmailsWithName.Items.Add("ripoff");
            lv_ExcludeEmailsWithName.Items.Add("scam");
            lv_ExcludeEmailsWithName.Items.Add("spam");
            lv_ExcludeEmailsWithName.Items.Add("spamcop");
            lv_ExcludeEmailsWithName.Items.Add("spammer");
            lv_ExcludeEmailsWithName.Items.Add("subscribe");
            lv_ExcludeEmailsWithName.Items.Add("support");
            lv_ExcludeEmailsWithName.Items.Add("unsubscribe");
            lv_ExcludeEmailsWithName.Items.Add("yourdomain");
            lv_ExcludeEmailsWithName.Items.Add("yourname");
            lv_ExcludeEmailsWithName.Items.Add("xxx");
            lv_ExcludeEmailsWithName.Items.Add("..");

            lv_ExcludeEmailsWithDomain.Items.Clear();
            lv_ExcludeEmailsWithDomain.Items.Add("antispam");            

            lv_ExcludeURLsWith.Items.Clear();
            lv_ExcludeURLsWith.Items.Add("addreply.php");
            lv_ExcludeURLsWith.Items.Add("editpost.php");
            lv_ExcludeURLsWith.Items.Add("formmail.php");
            lv_ExcludeURLsWith.Items.Add("javascript");
            lv_ExcludeURLsWith.Items.Add("pms.php");
            lv_ExcludeURLsWith.Items.Add("print.php");
            lv_ExcludeURLsWith.Items.Add("profil.php");
            lv_ExcludeURLsWith.Items.Add("reply.php");
            lv_ExcludeURLsWith.Items.Add("report.php");
            lv_ExcludeURLsWith.Items.Add("usercp.php");
            lv_ExcludeURLsWith.Items.Add("dictionary.reference.com");
            lv_ExcludeURLsWith.Items.Add("googleadservices.com");
            lv_ExcludeURLsWith.Items.Add("yahoo.com");
            lv_ExcludeURLsWith.Items.Add("msn.com");
            lv_ExcludeURLsWith.Items.Add("answers.com");
            lv_ExcludeURLsWith.Items.Add("altavista.com");
            lv_ExcludeURLsWith.Items.Add("doubleclick");
            lv_ExcludeURLsWith.Items.Add("scam");
            lv_ExcludeURLsWith.Items.Add("fraud");
            lv_ExcludeURLsWith.Items.Add(".dll");
            lv_ExcludeURLsWith.Items.Add(".exe");
            lv_ExcludeURLsWith.Items.Add(".bat");
            lv_ExcludeURLsWith.Items.Add(".ocx");
            lv_ExcludeURLsWith.Items.Add("w3.org"); 

            listOfEngines.Clear();

            listOfEngines.Add(new SearchEngine("Google.com", "https://www.google.com/search?q=", false));
            listOfEngines.Add(new SearchEngine("Google.ru", "https://www.google.ru/search?q=", false));
            listOfEngines.Add(new SearchEngine("Yandex.ru", "https://yandex.ru/yandsearch?text=", false));
            listOfEngines.Add(new SearchEngine("Bingo.com", "http://www.bing.com/search?q=", false));
            listOfEngines.Add(new SearchEngine("Forex Systems Ru", "http://forexsystems.ru/", true));
            listOfEngines.Add(new SearchEngine("Rambler.ru", "http://nova.rambler.ru/search?query=", false));
            listOfEngines.Add(new SearchEngine("Yahoo.com", "https://search.yahoo.com/search?p", false));
            listOfEngines.Add(new SearchEngine("Mail.ru", "http://go.mail.ru/search?q=", false));
            listOfEngines.Add(new SearchEngine("Gigablast.com", "https://www.gigablast.com/search?q=", false));
            listOfEngines.Add(new SearchEngine("Lycos.com", "http://search.lycos.com/web?q=", false));
            listOfEngines.Add(new SearchEngine("Сообщество Форекс Трейдеров", "http://ruforum.mt5.com/", false));
            listOfEngines.Add(new SearchEngine("FxClub.org", "http://forum.fxclub.org/", false));
            listOfEngines.Add(new SearchEngine("DailyFx.com", "http://www.dailyfx.com/forex_forum/", false));
            listOfEngines.Add(new SearchEngine("ForexFactory.com", "http://www.forexfactory.com/forum.php", false));
            listOfEngines.Add(new SearchEngine("Traders Territory", "http://tradersterritory.com/", false));
            listOfEngines.Add(new SearchEngine("Forex Pf Ru", "http://www.forexpf.ru/forum/", false));
            listOfEngines.Add(new SearchEngine("Forex Forum Ru", "http://forexforum.ru/", false));
            listOfEngines.Add(new SearchEngine("Fx-Trend", "http://forum.fx-trend.com/", false));
            listOfEngines.Add(new SearchEngine("Global-View.com", "http://www.global-view.com/forums/forum_ticker.html?f=1", false));
            listOfEngines.Add(new SearchEngine("Forex-TSD.com", "http://www.forex-tsd.com/", false));
            listOfEngines.Add(new SearchEngine("DonnaForex.com", "http://www.donnaforex.com/forum/", false));
            listOfEngines.Add(new SearchEngine("Dream Team Money", "http://www.dreamteammoney.com/index.php?showforum=48", false));
        
        

            listOfEngines = listOfEngines.OrderBy(o => o.getEngineName()).ToList();

            refreshSearchEnginesListView();

            checkBox1.Checked = true;
            checkBox2.Checked = false;
            checkBox3.Checked = false;

            ckbAutoRefresh1.Checked = true;
            ckbAutoRefresh2.Checked = true;
           
        }

        ///////////////////////////////////////////////////////////////////////
        // 
        // Choose Directory Path For Settings
        //

        private void button1_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                txtDefaultPathForSettings.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        //////////////////////////////////////////////////////////////////////
        //
        // Choose Directory Path For Campaigns
        //

        private void button5_Click(object sender, EventArgs e)
        {

            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                txtDefaultPathForCampaigns.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        //////////////////////////////////////////////////////////////////////
        //
        // Load XML settings
        //

        private void loadSettingsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            load_xml_settings();
        }

        ////////////////////////////////////////////////////////
        //
        // Load XML settings
        //

        private void openToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            load_xml_settings();
        }

        private void load_xml_settings()
        {
            try
            {
                openFileDialog1.RestoreDirectory = true;
                openFileDialog1.Title = "Load Settings XML";
                openFileDialog1.DefaultExt = "xml";
                openFileDialog1.Filter = "XML files (*.xml) |*.xml|All files (*.*)|*.*";
                openFileDialog1.FilterIndex = 1;
                openFileDialog1.InitialDirectory = txtDefaultPathForSettings.Text;
                openFileDialog1.FileName = "";

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string sStartupPath = openFileDialog1.FileName.ToString();

                    try
                    {
                        clsSValidator objclsSValidator =
                              new clsSValidator(sStartupPath,
                                                txtSettingsXSDFilePath.Text.ToString() +
                                                "\\" +
                                                txtSettingsXSDFileName.Text.ToString());

                        if (!objclsSValidator.ValidateXMLFile()) return;

                        XmlTextReader objXmlTextReader =
                              new XmlTextReader(sStartupPath);

                        string sName = "";

                        while (objXmlTextReader.Read())
                        {
                            switch (objXmlTextReader.NodeType)
                            {
                                case XmlNodeType.Element:
                                    sName = objXmlTextReader.Name;
                                    break;
                                case XmlNodeType.Text:
                                    switch (sName)
                                    {
                                        case "generalsettings.version":
                                            txtVersion.Text = objXmlTextReader.Value.ToString();
                                            break;
                                        case "generalsettings.apppath":
                                            txtAppPath.Text = objXmlTextReader.Value.ToString();
                                            break;
                                        case "generalsettings.settingsxsdfilename":
                                            txtSettingsXSDFileName.Text = objXmlTextReader.Value.ToString();
                                            break;
                                        case "generalsettings.settingsxsdpath":
                                            txtSettingsXSDFilePath.Text = objXmlTextReader.Value.ToString();
                                            break;
                                        case "generalsettings.settingsxmldeffilename":
                                            txtSettingsXMLDefaultFileName.Text = 
                                                objXmlTextReader.Value.ToString();
                                            break;
                                        case "generalsettings.enginefrequency":
                                            comboEngineFrequency.SelectedIndex = 
                                                Convert.ToInt32(objXmlTextReader.Value.ToString());
                                            break;
                                        case "generalsettings.degreeparallelism":
                                            comboDegreeParallelism.SelectedIndex = 
                                                Convert.ToInt32(objXmlTextReader.Value.ToString());
                                            break;
                                        case "generalsettings.maxbufsizeprepareurl":
                                            comboMaxBufferSzPrepareURL.SelectedIndex = 
                                                Convert.ToInt32(objXmlTextReader.Value.ToString());
                                            break;
                                        case "generalsettings.defaultpathforsettings":
                                            txtDefaultPathForSettings.Text =
                                                objXmlTextReader.Value.ToString();
                                            break;
                                        case "generalsettings.defaultpathforcampaings":
                                            txtDefaultPathForCampaigns.Text =
                                                objXmlTextReader.Value.ToString();
                                            break;
                                        case "generalsettings.maxsearchdepth":
                                            comboMaxSearchDepth.SelectedIndex =
                                                Convert.ToInt32(objXmlTextReader.Value.ToString());
                                            break;
                                        case "generalsettings.maxleadstoextract":
                                            comboMaxLeadsToExtract.SelectedIndex =
                                                Convert.ToInt32(objXmlTextReader.Value.ToString());
                                            break;
                                        case "generalsettings.nothingtoprocessattempts":
                                            cbxNothingToProcessAtmpts.SelectedIndex =
                                                Convert.ToInt32(objXmlTextReader.Value.ToString());
                                            break; 
                                        case "exportsettings.addheaderexportexcel":
                                            if (objXmlTextReader.Value=="True")
                                                chkboxAddHeaderExportExcel.Checked = true;
                                            else
                                                chkboxAddHeaderExportExcel.Checked = false;
                                            break;

                                        case "searchsettings.emailregularexpression":
                                            txtEmailRegEx.Text = objXmlTextReader.Value.ToString();
                                            break;

                                        case "searchsettings.phoneregularexpression":
                                            txtPhoneRegEx.Text = objXmlTextReader.Value.ToString();
                                            break;

                                        case "searchsettings.skyperegularexpression":
                                            txtSkypeRegEx.Text = objXmlTextReader.Value.ToString();
                                            break;

                                        case "searchsettings.maxemailsize":
                                            cbxMaxEmailSize.SelectedIndex =
                                                Convert.ToInt32(objXmlTextReader.Value.ToString());                                                
                                            break;

                                        case "searchsettings.maxphonesize":
                                            cbxMaxPhoneSize.SelectedIndex =
                                                Convert.ToInt32(objXmlTextReader.Value.ToString());
                                            break;

                                        case "searchsettings.maxskypesize":
                                            cbxMaxSkypeSize.SelectedIndex =
                                                Convert.ToInt32(objXmlTextReader.Value.ToString());
                                            break;

                                        case "searchsettings.maxwebpageize":
                                            cbxMaxWebPagesSize.SelectedIndex =
                                                Convert.ToInt32(objXmlTextReader.Value.ToString());
                                            break;

                                        case "searchsettings.excludeemailwithname":
                                            String strTmp1 =  objXmlTextReader.Value.ToString();
                                            char[] delimiterChars1 = { ',' };                                             
                                            string[] words1 = strTmp1.Split(delimiterChars1);
                                            lv_ExcludeEmailsWithName.Items.Clear();
                                            for (Int32 idx = 0; idx < words1.Count(); idx++)
                                            {                                               
                                                ListViewItem oItem = new ListViewItem();
                                                oItem.Text = words1[idx];                                                
                                                lv_ExcludeEmailsWithName.Items.Add(oItem);
                                            }
                                            break;

                                        case "searchsettings.excludeemailwithdomain":
                                            String strTmp2 =  objXmlTextReader.Value.ToString();
                                            char[] delimiterChars2 = { ',' };                                             
                                            string[] words2 = strTmp2.Split(delimiterChars2);
                                            lv_ExcludeEmailsWithDomain.Items.Clear();
                                            for (Int32 idx = 0; idx < words2.Count(); idx++)
                                            {
                                                ListViewItem oItem = new ListViewItem();
                                                oItem.Text = words2[idx];
                                                lv_ExcludeEmailsWithDomain.Items.Add(oItem);                                                
                                            }
                                            break;

                                        case "searchsettings.excludeurlwith":
                                            String strTmp3 = objXmlTextReader.Value.ToString();
                                            char[] delimiterChars3 = { ',' };
                                            string[] words3 = strTmp3.Split(delimiterChars3);
                                            lv_ExcludeURLsWith.Items.Clear();
                                            for (Int32 idx = 0; idx < words3.Count(); idx++)
                                            {
                                                ListViewItem oItem = new ListViewItem();
                                                oItem.Text = words3[idx];
                                                lv_ExcludeURLsWith.Items.Add(oItem); 
                                            }
                                            break;

                                        case "searchsettings.autosaveresults":
                                            cbxAutoSaveResults.SelectedIndex =
                                                Convert.ToInt32(objXmlTextReader.Value.ToString());
                                            break;

                                        case "startsearchwithurls.searchengines":
                                            String strTmp4 = objXmlTextReader.Value.ToString();
                                            char[] delimiterChars4 = { '^' };
                                            string[] words4 = strTmp4.Split(delimiterChars4);
                                            listViewSearchEngines.Items.Clear();
                                            for (Int32 idx = 0; idx < words4.Count(); idx++)
                                            {

                                                String strTmp5 = words4[idx];
                                                char[] delimiterChars5 = { '|' };
                                                string[] words5 = strTmp5.Split(delimiterChars5);

                                                ListViewItem oItem = new ListViewItem();

                                                if (words5[0] == "True")
                                                    oItem.Checked = true;
                                                else
                                                    oItem.Checked = false;

                                                oItem.Text = words5[1];
                                                oItem.SubItems.Add(words5[2]);
                                                listViewSearchEngines.Items.Add(oItem);
                                            }
                                            break;

                                        case "extractbykeywords.filterseeemails":
                                            checkBox1.Checked =
                                                Convert.ToBoolean(objXmlTextReader.Value.ToString());
                                            break;

                                        case "extractbykeywords.filterseephones":
                                            checkBox2.Checked =
                                                Convert.ToBoolean(objXmlTextReader.Value.ToString());
                                            break;

                                        case "extractbykeywords.filterseeskypes":
                                            checkBox3.Checked =
                                                Convert.ToBoolean(objXmlTextReader.Value.ToString());
                                            break;

                                        case "extractbykeywords.keywordoneverypage":
                                            checkBox5.Checked =
                                                Convert.ToBoolean(objXmlTextReader.Value.ToString());
                                            break;

                                        case "extractbykeywords.keywordappearinurl":
                                            checkBox4.Checked =
                                                Convert.ToBoolean(objXmlTextReader.Value.ToString());
                                            break;
                                    }
                                    break;
                            }
                        } // while

                        objXmlTextReader.Close();
                        
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        private void save_as_xml_settings()
        {
            try
            {
                saveFileDialog1.RestoreDirectory = true;
                saveFileDialog1.Title = "Save Settings XML";
                saveFileDialog1.DefaultExt = "xml";
                saveFileDialog1.Filter = "XML files (*.xml) |*.xml|All files (*.*)|*.*";
                saveFileDialog1.FilterIndex = 1;
                saveFileDialog1.RestoreDirectory = true;
                saveFileDialog1.InitialDirectory = txtDefaultPathForSettings.Text;

                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string sStartupPath = saveFileDialog1.FileName.ToString();
                    XmlTextWriter objXmlTextWriter =
                            new XmlTextWriter(sStartupPath, null);
                    objXmlTextWriter.Formatting = Formatting.Indented;
                    objXmlTextWriter.WriteStartDocument();

                    objXmlTextWriter.WriteStartElement("LeadsExtractorSettings");

                    objXmlTextWriter.WriteStartElement("generalsettings.version");
                    objXmlTextWriter.WriteString(txtVersion.Text);
                    objXmlTextWriter.WriteEndElement();

                    objXmlTextWriter.WriteStartElement("generalsettings.apppath");
                    objXmlTextWriter.WriteString(txtAppPath.Text);
                    objXmlTextWriter.WriteEndElement();

                    objXmlTextWriter.WriteStartElement("generalsettings.settingsxsdfilename");
                    objXmlTextWriter.WriteString(txtSettingsXSDFileName.Text);
                    objXmlTextWriter.WriteEndElement();

                    objXmlTextWriter.WriteStartElement("generalsettings.settingsxsdpath");
                    objXmlTextWriter.WriteString(txtSettingsXSDFilePath.Text);
                    objXmlTextWriter.WriteEndElement();

                    objXmlTextWriter.WriteStartElement("generalsettings.settingsxmldeffilename");
                    objXmlTextWriter.WriteString(txtSettingsXMLDefaultFileName.Text);
                    objXmlTextWriter.WriteEndElement();

                    objXmlTextWriter.WriteStartElement("generalsettings.enginefrequency");
                    objXmlTextWriter.WriteString(comboEngineFrequency.SelectedIndex.ToString());
                    objXmlTextWriter.WriteEndElement();

                    objXmlTextWriter.WriteStartElement("generalsettings.degreeparallelism");
                    objXmlTextWriter.WriteString(comboDegreeParallelism.SelectedIndex.ToString());
                    objXmlTextWriter.WriteEndElement();

                    objXmlTextWriter.WriteStartElement("generalsettings.maxbufsizeprepareurl");
                    objXmlTextWriter.WriteString(comboMaxBufferSzPrepareURL.SelectedIndex.ToString());
                    objXmlTextWriter.WriteEndElement();

                    objXmlTextWriter.WriteStartElement("generalsettings.defaultpathforsettings");
                    objXmlTextWriter.WriteString(txtDefaultPathForSettings.Text);
                    objXmlTextWriter.WriteEndElement();

                    objXmlTextWriter.WriteStartElement("generalsettings.defaultpathforcampaings");
                    objXmlTextWriter.WriteString(txtDefaultPathForCampaigns.Text);
                    objXmlTextWriter.WriteEndElement();

                    objXmlTextWriter.WriteStartElement("generalsettings.maxsearchdepth");
                    objXmlTextWriter.WriteString(comboMaxSearchDepth.SelectedIndex.ToString());
                    objXmlTextWriter.WriteEndElement();

                    objXmlTextWriter.WriteStartElement("generalsettings.maxleadstoextract");
                    objXmlTextWriter.WriteString(comboMaxLeadsToExtract.SelectedIndex.ToString());
                    objXmlTextWriter.WriteEndElement();

                    objXmlTextWriter.WriteStartElement("generalsettings.nothingtoprocessattempts");
                    objXmlTextWriter.WriteString(cbxNothingToProcessAtmpts.SelectedIndex.ToString());
                    objXmlTextWriter.WriteEndElement();

                    objXmlTextWriter.WriteStartElement("exportsettings.addheaderexportexcel");
                    objXmlTextWriter.WriteString(chkboxAddHeaderExportExcel.Checked.ToString());
                    objXmlTextWriter.WriteEndElement();

                    objXmlTextWriter.WriteStartElement("searchsettings.emailregularexpression");
                    objXmlTextWriter.WriteString(txtEmailRegEx.Text);
                    objXmlTextWriter.WriteEndElement();

                    objXmlTextWriter.WriteStartElement("searchsettings.phoneregularexpression");
                    objXmlTextWriter.WriteString(txtPhoneRegEx.Text);
                    objXmlTextWriter.WriteEndElement();

                    objXmlTextWriter.WriteStartElement("searchsettings.skyperegularexpression");
                    objXmlTextWriter.WriteString(txtSkypeRegEx.Text);
                    objXmlTextWriter.WriteEndElement();

                    objXmlTextWriter.WriteStartElement("searchsettings.maxemailsize");
                    objXmlTextWriter.WriteString(cbxMaxEmailSize.SelectedIndex.ToString());
                    objXmlTextWriter.WriteEndElement();

                    objXmlTextWriter.WriteStartElement("searchsettings.maxphonesize");
                    objXmlTextWriter.WriteString(cbxMaxPhoneSize.SelectedIndex.ToString());
                    objXmlTextWriter.WriteEndElement();

                    objXmlTextWriter.WriteStartElement("searchsettings.maxskypesize");
                    objXmlTextWriter.WriteString(cbxMaxSkypeSize.SelectedIndex.ToString());
                    objXmlTextWriter.WriteEndElement();

                    objXmlTextWriter.WriteStartElement("searchsettings.maxwebpageize");
                    objXmlTextWriter.WriteString(cbxMaxWebPagesSize.SelectedIndex.ToString());
                    objXmlTextWriter.WriteEndElement();

                    String strTmp = "";
                    for (Int32 idx = 0; idx < lv_ExcludeEmailsWithName.Items.Count; idx++)
                    {
                        strTmp = strTmp + lv_ExcludeEmailsWithName.Items[idx].Text;

                        if (idx != lv_ExcludeEmailsWithName.Items.Count - 1)
                            strTmp = strTmp + ",";
                    }
                    objXmlTextWriter.WriteStartElement("searchsettings.excludeemailwithname");
                    objXmlTextWriter.WriteString(strTmp);
                    objXmlTextWriter.WriteEndElement();

                    strTmp = "";
                    for (Int32 idx = 0; idx < lv_ExcludeEmailsWithDomain.Items.Count; idx++)
                    {
                        strTmp = strTmp + lv_ExcludeEmailsWithDomain.Items[idx].Text;

                        if (idx != lv_ExcludeEmailsWithDomain.Items.Count - 1)
                            strTmp = strTmp + ",";
                    }
                    objXmlTextWriter.WriteStartElement("searchsettings.excludeemailwithdomain");
                    objXmlTextWriter.WriteString(strTmp);
                    objXmlTextWriter.WriteEndElement();

                    strTmp = "";
                    for (Int32 idx = 0; idx < lv_ExcludeURLsWith.Items.Count; idx++)
                    {
                        strTmp = strTmp + lv_ExcludeURLsWith.Items[idx].Text;

                        if (idx != lv_ExcludeURLsWith.Items.Count - 1)
                            strTmp = strTmp + ",";
                    }
                    objXmlTextWriter.WriteStartElement("searchsettings.excludeurlwith");
                    objXmlTextWriter.WriteString(strTmp);
                    objXmlTextWriter.WriteEndElement();


                    objXmlTextWriter.WriteStartElement("searchsettings.autosaveresults");
                    objXmlTextWriter.WriteString(cbxAutoSaveResults.SelectedIndex.ToString());
                    objXmlTextWriter.WriteEndElement();

                    strTmp = "";
                    for (Int32 idx = 0; idx < listViewSearchEngines.Items.Count; idx++)
                    {
                        strTmp = strTmp + listViewSearchEngines.Items[idx].Checked.ToString() +
                                "|" + listViewSearchEngines.Items[idx].Text +
                                "|" + listViewSearchEngines.Items[idx].SubItems[1].Text;

                        if (idx != listViewSearchEngines.Items.Count - 1)
                            strTmp = strTmp + "^";
                    }
                    objXmlTextWriter.WriteStartElement("startsearchwithurls.searchengines");
                    objXmlTextWriter.WriteString(strTmp);
                    objXmlTextWriter.WriteEndElement();


                    objXmlTextWriter.WriteStartElement("extractbykeywords.filterseeemails");
                    objXmlTextWriter.WriteString(checkBox1.Checked.ToString());
                    objXmlTextWriter.WriteEndElement();

                    objXmlTextWriter.WriteStartElement("extractbykeywords.filterseephones");
                    objXmlTextWriter.WriteString(checkBox2.Checked.ToString());
                    objXmlTextWriter.WriteEndElement();

                    objXmlTextWriter.WriteStartElement("extractbykeywords.filterseeskypes");
                    objXmlTextWriter.WriteString(checkBox3.Checked.ToString());
                    objXmlTextWriter.WriteEndElement();

                    objXmlTextWriter.WriteStartElement("extractbykeywords.keywordoneverypage");
                    objXmlTextWriter.WriteString(checkBox5.Checked.ToString());
                    objXmlTextWriter.WriteEndElement();

                    objXmlTextWriter.WriteStartElement("extractbykeywords.keywordappearinurl");
                    objXmlTextWriter.WriteString(checkBox4.Checked.ToString());
                    objXmlTextWriter.WriteEndElement();

                    objXmlTextWriter.WriteStartElement("extractbykeywords.autorefresh");
                    objXmlTextWriter.WriteString(ckbAutoRefresh1.Checked.ToString());
                    objXmlTextWriter.WriteEndElement();

                    objXmlTextWriter.WriteStartElement("processurlsprogress.autorefresh");
                    objXmlTextWriter.WriteString(ckbAutoRefresh2.Checked.ToString());
                    objXmlTextWriter.WriteEndElement();

                    objXmlTextWriter.WriteEndDocument();
                    objXmlTextWriter.Flush();
                    objXmlTextWriter.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        /////////////////////////////////////////////////////////////////////
        //
        // load xml settings on start
        //

        private void load_xml_settings_on_start()
        {

            try
            {

                restoreDefaultSettings();

                String strAppPath = Directory.GetCurrentDirectory().ToString();

                string sStartupPath = strAppPath + "\\" + "LeadsExtractorSettings.xml";

                string sXSDfile = strAppPath + "\\" + "LeadsExtractorSettings.xsd";  

                if (File.Exists(sStartupPath) == false || File.Exists(sXSDfile) == false)
                {
                    restoreDefaultSettings();
                }
                else
                {
                    clsSValidator objclsSValidator =
                          new clsSValidator(sStartupPath,
                                            sXSDfile);

                    if (!objclsSValidator.ValidateXMLFile())
                    {
                        restoreDefaultSettings();
                    }
                    else
                    {

                        XmlTextReader objXmlTextReader =
                              new XmlTextReader(sStartupPath);

                        string sName = "";

                        while (objXmlTextReader.Read())
                        {
                            switch (objXmlTextReader.NodeType)
                            {
                                case XmlNodeType.Element:
                                    sName = objXmlTextReader.Name;
                                    break;
                                case XmlNodeType.Text:
                                    switch (sName)
                                    {
                                        case "generalsettings.version":
                                            txtVersion.Text = objXmlTextReader.Value.ToString();
                                            break;
                                        case "generalsettings.apppath":
                                            txtAppPath.Text = objXmlTextReader.Value.ToString();
                                            break;
                                        case "generalsettings.settingsxsdfilename":
                                            txtSettingsXSDFileName.Text = objXmlTextReader.Value.ToString();
                                            break;
                                        case "generalsettings.settingsxsdpath":
                                            txtSettingsXSDFilePath.Text = objXmlTextReader.Value.ToString();
                                            break;
                                        case "generalsettings.settingsxmldeffilename":
                                            txtSettingsXMLDefaultFileName.Text =
                                                objXmlTextReader.Value.ToString();
                                            break;
                                        case "generalsettings.enginefrequency":
                                            comboEngineFrequency.SelectedIndex =
                                                Convert.ToInt32(objXmlTextReader.Value.ToString());
                                            break;
                                        case "generalsettings.degreeparallelism":
                                            comboDegreeParallelism.SelectedIndex =
                                                Convert.ToInt32(objXmlTextReader.Value.ToString());
                                            break;
                                        case "generalsettings.maxbufsizeprepareurl":
                                            comboMaxBufferSzPrepareURL.SelectedIndex =
                                                Convert.ToInt32(objXmlTextReader.Value.ToString());
                                            break;
                                        case "generalsettings.defaultpathforsettings":
                                            txtDefaultPathForSettings.Text =
                                                objXmlTextReader.Value.ToString();
                                            break;
                                        case "generalsettings.defaultpathforcampaings":
                                            txtDefaultPathForCampaigns.Text =
                                                objXmlTextReader.Value.ToString();
                                            break;
                                        case "generalsettings.maxsearchdepth":
                                            comboMaxSearchDepth.SelectedIndex =
                                                Convert.ToInt32(objXmlTextReader.Value.ToString());
                                            break;
                                        case "generalsettings.maxleadstoextract":
                                            comboMaxLeadsToExtract.SelectedIndex =
                                                Convert.ToInt32(objXmlTextReader.Value.ToString());
                                            break;
                                        case "generalsettings.nothingtoprocessattempts":
                                            cbxNothingToProcessAtmpts.SelectedIndex =
                                                Convert.ToInt32(objXmlTextReader.Value.ToString());
                                            break;
                                        case "exportsettings.addheaderexportexcel":
                                            if (objXmlTextReader.Value == "True")
                                                chkboxAddHeaderExportExcel.Checked = true;
                                            else
                                                chkboxAddHeaderExportExcel.Checked = false;
                                            break;

                                        case "searchsettings.emailregularexpression":
                                            txtEmailRegEx.Text = objXmlTextReader.Value.ToString();
                                            break;

                                        case "searchsettings.phoneregularexpression":
                                            txtPhoneRegEx.Text = objXmlTextReader.Value.ToString();
                                            break;

                                        case "searchsettings.skyperegularexpression":
                                            txtSkypeRegEx.Text = objXmlTextReader.Value.ToString();
                                            break;

                                        case "searchsettings.maxemailsize":
                                            cbxMaxEmailSize.SelectedIndex =
                                                Convert.ToInt32(objXmlTextReader.Value.ToString());
                                            break;

                                        case "searchsettings.maxphonesize":
                                            cbxMaxPhoneSize.SelectedIndex =
                                                Convert.ToInt32(objXmlTextReader.Value.ToString());
                                            break;

                                        case "searchsettings.maxskypesize":
                                            cbxMaxSkypeSize.SelectedIndex =
                                                Convert.ToInt32(objXmlTextReader.Value.ToString());
                                            break;

                                        case "searchsettings.maxwebpageize":
                                            cbxMaxWebPagesSize.SelectedIndex =
                                                Convert.ToInt32(objXmlTextReader.Value.ToString());
                                            break;

                                        case "searchsettings.excludeemailwithname":
                                            String strTmp1 = objXmlTextReader.Value.ToString();
                                            char[] delimiterChars1 = { ',' };
                                            string[] words1 = strTmp1.Split(delimiterChars1);
                                            lv_ExcludeEmailsWithName.Items.Clear();
                                            for (Int32 idx = 0; idx < words1.Count(); idx++)
                                            {
                                                ListViewItem oItem = new ListViewItem();
                                                oItem.Text = words1[idx];
                                                lv_ExcludeEmailsWithName.Items.Add(oItem);
                                            }
                                            break;

                                        case "searchsettings.excludeemailwithdomain":
                                            String strTmp2 = objXmlTextReader.Value.ToString();
                                            char[] delimiterChars2 = { ',' };
                                            string[] words2 = strTmp2.Split(delimiterChars2);
                                            lv_ExcludeEmailsWithDomain.Items.Clear();
                                            for (Int32 idx = 0; idx < words2.Count(); idx++)
                                            {
                                                ListViewItem oItem = new ListViewItem();
                                                oItem.Text = words2[idx];
                                                lv_ExcludeEmailsWithDomain.Items.Add(oItem);
                                            }
                                            break;

                                        case "searchsettings.excludeurlwith":
                                            String strTmp3 = objXmlTextReader.Value.ToString();
                                            char[] delimiterChars3 = { ',' };
                                            string[] words3 = strTmp3.Split(delimiterChars3);
                                            lv_ExcludeURLsWith.Items.Clear();
                                            for (Int32 idx = 0; idx < words3.Count(); idx++)
                                            {
                                                ListViewItem oItem = new ListViewItem();
                                                oItem.Text = words3[idx];
                                                lv_ExcludeURLsWith.Items.Add(oItem);
                                            }
                                            break;

                                        case "searchsettings.autosaveresults":
                                            cbxAutoSaveResults.SelectedIndex =
                                                Convert.ToInt32(objXmlTextReader.Value.ToString());
                                            break;

                                        case "startsearchwithurls.searchengines":
                                            String strTmp4 = objXmlTextReader.Value.ToString();
                                            char[] delimiterChars4 = { '^' };
                                            string[] words4 = strTmp4.Split(delimiterChars4);
                                            listViewSearchEngines.Items.Clear();
                                            for (Int32 idx = 0; idx < words4.Count(); idx++)
                                            {

                                                String strTmp5 = words4[idx];
                                                char[] delimiterChars5 = { '|' };
                                                string[] words5 = strTmp5.Split(delimiterChars5);

                                                ListViewItem oItem = new ListViewItem();

                                                if (words5[0] == "True")
                                                    oItem.Checked = true;
                                                else
                                                    oItem.Checked = false;

                                                oItem.Text = words5[1];
                                                oItem.SubItems.Add(words5[2]);
                                                listViewSearchEngines.Items.Add(oItem);
                                            }
                                            break;

                                        case "extractbykeywords.filterseeemails":
                                            checkBox1.Checked =
                                                Convert.ToBoolean(objXmlTextReader.Value.ToString());
                                            break;

                                        case "extractbykeywords.filterseephones":
                                            checkBox2.Checked =
                                                Convert.ToBoolean(objXmlTextReader.Value.ToString());
                                            break;

                                        case "extractbykeywords.filterseeskypes":
                                            checkBox3.Checked =
                                                Convert.ToBoolean(objXmlTextReader.Value.ToString());
                                            break;

                                        case "extractbykeywords.keywordoneverypage":
                                            checkBox5.Checked =
                                                Convert.ToBoolean(objXmlTextReader.Value.ToString());
                                            break;

                                        case "extractbykeywords.keywordappearinurl":
                                            checkBox4.Checked =
                                                Convert.ToBoolean(objXmlTextReader.Value.ToString());
                                            break;

                                        case "extractbykeywords.autorefresh":
                                            ckbAutoRefresh1.Checked =
                                                Convert.ToBoolean(objXmlTextReader.Value.ToString());
                                            break;

                                        case "processurlsprogress.autorefresh":
                                            ckbAutoRefresh2.Checked =
                                                Convert.ToBoolean(objXmlTextReader.Value.ToString());
                                            break;
                                    }
                                    break;
                            }
                        } // while

                        objXmlTextReader.Close();

                    }

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


        }

        ////////////////////////////////////////////////////////////////////
        // 
        // save xml settings
        //

        private void saveToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            save_xml_settings();                
        }

        private void saveAsToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            save_as_xml_settings();
        }

        private void save_xml_settings()
        {
            try
            {

                string sStartupPath = txtDefaultPathForSettings.Text.ToString() +
                                                "\\" +
                                                txtSettingsXMLDefaultFileName.Text.ToString();

                XmlTextWriter objXmlTextWriter =
                        new XmlTextWriter(sStartupPath, null);
                
                objXmlTextWriter.Formatting = Formatting.Indented;
                objXmlTextWriter.WriteStartDocument();

                objXmlTextWriter.WriteStartElement("LeadsExtractorSettings");

                objXmlTextWriter.WriteStartElement("generalsettings.version");
                objXmlTextWriter.WriteString(txtVersion.Text);
                objXmlTextWriter.WriteEndElement();

                objXmlTextWriter.WriteStartElement("generalsettings.apppath");
                objXmlTextWriter.WriteString(txtAppPath.Text);
                objXmlTextWriter.WriteEndElement();

                objXmlTextWriter.WriteStartElement("generalsettings.settingsxsdfilename");
                objXmlTextWriter.WriteString(txtSettingsXSDFileName.Text);
                objXmlTextWriter.WriteEndElement();

                objXmlTextWriter.WriteStartElement("generalsettings.settingsxsdpath");
                objXmlTextWriter.WriteString(txtSettingsXSDFilePath.Text);
                objXmlTextWriter.WriteEndElement();

                objXmlTextWriter.WriteStartElement("generalsettings.settingsxmldeffilename");
                objXmlTextWriter.WriteString(txtSettingsXMLDefaultFileName.Text);
                objXmlTextWriter.WriteEndElement();

                objXmlTextWriter.WriteStartElement("generalsettings.enginefrequency");
                objXmlTextWriter.WriteString(comboEngineFrequency.SelectedIndex.ToString());
                objXmlTextWriter.WriteEndElement();

                objXmlTextWriter.WriteStartElement("generalsettings.degreeparallelism");
                objXmlTextWriter.WriteString(comboDegreeParallelism.SelectedIndex.ToString());
                objXmlTextWriter.WriteEndElement();

                objXmlTextWriter.WriteStartElement("generalsettings.maxbufsizeprepareurl");
                objXmlTextWriter.WriteString(comboMaxBufferSzPrepareURL.SelectedIndex.ToString());
                objXmlTextWriter.WriteEndElement();

                objXmlTextWriter.WriteStartElement("generalsettings.defaultpathforsettings");
                objXmlTextWriter.WriteString(txtDefaultPathForSettings.Text);
                objXmlTextWriter.WriteEndElement();

                objXmlTextWriter.WriteStartElement("generalsettings.defaultpathforcampaings");
                objXmlTextWriter.WriteString(txtDefaultPathForCampaigns.Text);
                objXmlTextWriter.WriteEndElement();

                objXmlTextWriter.WriteStartElement("generalsettings.maxsearchdepth");
                objXmlTextWriter.WriteString(comboMaxSearchDepth.SelectedIndex.ToString());
                objXmlTextWriter.WriteEndElement();

                objXmlTextWriter.WriteStartElement("generalsettings.maxleadstoextract");
                objXmlTextWriter.WriteString(comboMaxLeadsToExtract.SelectedIndex.ToString());
                objXmlTextWriter.WriteEndElement();

                objXmlTextWriter.WriteStartElement("generalsettings.nothingtoprocessattempts");
                objXmlTextWriter.WriteString(cbxNothingToProcessAtmpts.SelectedIndex.ToString());
                objXmlTextWriter.WriteEndElement();

                objXmlTextWriter.WriteStartElement("exportsettings.addheaderexportexcel");
                objXmlTextWriter.WriteString(chkboxAddHeaderExportExcel.Checked.ToString());
                objXmlTextWriter.WriteEndElement();

                objXmlTextWriter.WriteStartElement("searchsettings.emailregularexpression");
                objXmlTextWriter.WriteString(txtEmailRegEx.Text);
                objXmlTextWriter.WriteEndElement();

                objXmlTextWriter.WriteStartElement("searchsettings.phoneregularexpression");
                objXmlTextWriter.WriteString(txtPhoneRegEx.Text);
                objXmlTextWriter.WriteEndElement();

                objXmlTextWriter.WriteStartElement("searchsettings.skyperegularexpression");
                objXmlTextWriter.WriteString(txtSkypeRegEx.Text);
                objXmlTextWriter.WriteEndElement();

                objXmlTextWriter.WriteStartElement("searchsettings.maxemailsize");
                objXmlTextWriter.WriteString(cbxMaxEmailSize.SelectedIndex.ToString());
                objXmlTextWriter.WriteEndElement();

                objXmlTextWriter.WriteStartElement("searchsettings.maxphonesize");
                objXmlTextWriter.WriteString(cbxMaxPhoneSize.SelectedIndex.ToString());
                objXmlTextWriter.WriteEndElement();

                objXmlTextWriter.WriteStartElement("searchsettings.maxskypesize");
                objXmlTextWriter.WriteString(cbxMaxSkypeSize.SelectedIndex.ToString());
                objXmlTextWriter.WriteEndElement();

                objXmlTextWriter.WriteStartElement("searchsettings.maxwebpageize");
                objXmlTextWriter.WriteString(cbxMaxWebPagesSize.SelectedIndex.ToString());
                objXmlTextWriter.WriteEndElement();

                String strTmp="";
                for (Int32 idx = 0; idx < lv_ExcludeEmailsWithName.Items.Count; idx++)
                {
                    strTmp = strTmp + lv_ExcludeEmailsWithName.Items[idx].Text;

                    if (idx != lv_ExcludeEmailsWithName.Items.Count - 1)
                        strTmp = strTmp + ",";
                }
                objXmlTextWriter.WriteStartElement("searchsettings.excludeemailwithname");
                objXmlTextWriter.WriteString(strTmp);
                objXmlTextWriter.WriteEndElement();

                strTmp = "";
                for (Int32 idx = 0; idx < lv_ExcludeEmailsWithDomain.Items.Count; idx++)
                {
                    strTmp = strTmp + lv_ExcludeEmailsWithDomain.Items[idx].Text;

                    if (idx != lv_ExcludeEmailsWithDomain.Items.Count - 1)
                        strTmp = strTmp + ",";
                }
                objXmlTextWriter.WriteStartElement("searchsettings.excludeemailwithdomain");
                objXmlTextWriter.WriteString(strTmp);
                objXmlTextWriter.WriteEndElement();

                strTmp = "";
                for (Int32 idx = 0; idx < lv_ExcludeURLsWith.Items.Count; idx++)
                {
                    strTmp = strTmp + lv_ExcludeURLsWith.Items[idx].Text;

                    if (idx != lv_ExcludeURLsWith.Items.Count - 1)
                        strTmp = strTmp + ",";
                }
                objXmlTextWriter.WriteStartElement("searchsettings.excludeurlwith");
                objXmlTextWriter.WriteString(strTmp);
                objXmlTextWriter.WriteEndElement();


                objXmlTextWriter.WriteStartElement("searchsettings.autosaveresults");
                objXmlTextWriter.WriteString(cbxAutoSaveResults.SelectedIndex.ToString());
                objXmlTextWriter.WriteEndElement();
                
                strTmp = "";
                for (Int32 idx = 0; idx < listViewSearchEngines.Items.Count; idx++)
                {
                    strTmp = strTmp + listViewSearchEngines.Items[idx].Checked.ToString() +
                            "|" + listViewSearchEngines.Items[idx].Text +                            
                            "|" + listViewSearchEngines.Items[idx].SubItems[1].Text;

                    if (idx != listViewSearchEngines.Items.Count - 1)
                        strTmp = strTmp + "^";
                }
                objXmlTextWriter.WriteStartElement("startsearchwithurls.searchengines");
                objXmlTextWriter.WriteString(strTmp);
                objXmlTextWriter.WriteEndElement();

                objXmlTextWriter.WriteStartElement("extractbykeywords.filterseeemails");
                objXmlTextWriter.WriteString(checkBox1.Checked.ToString());
                objXmlTextWriter.WriteEndElement();

                objXmlTextWriter.WriteStartElement("extractbykeywords.filterseephones");
                objXmlTextWriter.WriteString(checkBox2.Checked.ToString());
                objXmlTextWriter.WriteEndElement();

                objXmlTextWriter.WriteStartElement("extractbykeywords.filterseeskypes");
                objXmlTextWriter.WriteString(checkBox3.Checked.ToString());
                objXmlTextWriter.WriteEndElement();

                objXmlTextWriter.WriteStartElement("extractbykeywords.keywordoneverypage");
                objXmlTextWriter.WriteString(checkBox5.Checked.ToString());
                objXmlTextWriter.WriteEndElement();

                objXmlTextWriter.WriteStartElement("extractbykeywords.keywordappearinurl");
                objXmlTextWriter.WriteString(checkBox4.Checked.ToString());
                objXmlTextWriter.WriteEndElement();

                objXmlTextWriter.WriteStartElement("extractbykeywords.autorefresh");
                objXmlTextWriter.WriteString(ckbAutoRefresh1.Checked.ToString());
                objXmlTextWriter.WriteEndElement();

                objXmlTextWriter.WriteStartElement("processurlsprogress.autorefresh");
                objXmlTextWriter.WriteString(ckbAutoRefresh2.Checked.ToString());
                objXmlTextWriter.WriteEndElement();

                objXmlTextWriter.WriteEndDocument();
                objXmlTextWriter.Flush();
                objXmlTextWriter.Close();
               
               
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void comboEngineFrequency_SelectedIndexChanged(object sender, EventArgs e)
        {
            timer1.Interval = Convert.ToInt32(comboEngineFrequency.Text);
        }

        ////////////////////////////////////////////////////////////
        //
        // copy selected leads to clipboard
        //

        private void copySelectedToClipboardToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // copy selected to clipboard

            Clipboard.Clear();
            StringBuilder buffer = new StringBuilder();

            // Setup the columns

            for (int i = 0; i < this.listView3.Columns.Count; i++)
            {
                buffer.Append(this.listView3.Columns[i].Text);
                buffer.Append("\t");
            }
            buffer.Append(Environment.NewLine);

            // Build the data row by row

            for (int i = 0; i < this.listView3.Items.Count; i++)
            {

                if (this.listView3.Items[i].Selected == true)
                {
                    for (int j = 0; j < this.listView3.Columns.Count; j++)
                    {
                        buffer.Append(this.listView3.Items[i].SubItems[j].Text);
                        buffer.Append("\t");
                    }
                    buffer.Append(Environment.NewLine);
                }
            }

            Clipboard.SetText(buffer.ToString());

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            // show emails check
            refreshLeadsList();
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            // show phones check
            refreshLeadsList();
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            // show skypes check
            refreshLeadsList();
        }

        private void refreshLeadsList()
        {
            
                // listview3 - leads

                listView3.Items.Clear();

                for (int idx = 0; idx < listOfLeads.Count; idx++)
                {

                    if (
                        (checkBox1.Checked == true) && (listOfLeads[idx].getLeadType() == LeadType.Email)
                        ||
                        (checkBox2.Checked == true) && (listOfLeads[idx].getLeadType() == LeadType.Phone)
                        ||
                        (checkBox3.Checked == true) && (listOfLeads[idx].getLeadType() == LeadType.Skype)
                        )
                    {

                        ListViewItem oItem = new ListViewItem();

                        oItem.Text = listOfLeads[idx].getLeadId().ToString();
                        oItem.SubItems.Add(listOfLeads[idx].getLeadTimeStamp().ToString(@"dd/MM/yyyy HH:mm:ss"));
                        oItem.SubItems.Add(listOfLeads[idx].getLeadType().ToString());
                        oItem.SubItems.Add(listOfLeads[idx].getLeadContent());
                        oItem.SubItems.Add(listOfLeads[idx].getLeadURL());
                        listView3.Items.Add(oItem);

                    }

                    // ensure visible
                    if (listView3.Items.Count - 1 > 0)
                    {
                        listView3.EnsureVisible(listView3.Items.Count - 1);
                    }
                   
                } //for
                  
        }

        private void comboDegreeParallelism_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            // refresh ui  

            if (IsRefreshLeadsList == false) return;

            IsRefreshLeadsList = false;

            removeDuplicateLeads();

            Int32 nOfLeads = 0;

            if (ckbAutoRefresh2.Checked == true)
            {

                // refresh listView2
                if (listOfURLsProcessed.Count == listView2.Items.Count)
                {
                    // nothing changed
                }
                else
                {

                    listView2.BeginUpdate();
                    try
                    {

                        listView2.Items.Clear();

                        for (int idx = 0; idx < listOfURLsProcessed.Count; idx++)
                        {
                            ListViewItem oItem = new ListViewItem();
                            oItem.Text = listOfURLsProcessed[idx].getURLProcessed();
                            oItem.SubItems.Add(listOfURLsProcessed[idx].getnOfEmails().ToString());
                            oItem.SubItems.Add(listOfURLsProcessed[idx].getnOfPhones().ToString());
                            oItem.SubItems.Add(listOfURLsProcessed[idx].getnOfSkypes().ToString());
                            oItem.SubItems.Add(listOfURLsProcessed[idx].getnDepthLevel().ToString());
                            oItem.SubItems.Add(listOfURLsProcessed[idx].getSearchParent().ToString());
                            oItem.SubItems.Add(listOfURLsProcessed[idx].getSearchHistoryChain().ToString());
                            listView2.Items.Add(oItem);

                        } // for

                        // ensure visible
                        if (listView2.Items.Count - 1 > 0)
                        {
                            listView2.EnsureVisible(listView2.Items.Count - 1);
                        }

                    }
                    finally
                    {
                        listView2.EndUpdate();
                    }

                }

            }

            if (ckbAutoRefresh1.Checked == true)
            {
                int nCheckQty = 0;

                for (int idx = 0; idx < listOfLeads.Count; idx++)
                {
                    if (
                        (checkBox1.Checked == true) && (listOfLeads[idx].getLeadType() == LeadType.Email)
                        ||
                        (checkBox2.Checked == true) && (listOfLeads[idx].getLeadType() == LeadType.Phone)
                        ||
                        (checkBox3.Checked == true) && (listOfLeads[idx].getLeadType() == LeadType.Skype)
                        )
                    {
                        nCheckQty++;
                    }

                } // for

                // refresh listView3

                if (listView3.Items.Count == nCheckQty)
                {
                    // nothing to do
                }
                else
                {

                    listView3.BeginUpdate();
                    try
                    {
                        listView3.Items.Clear();

                        for (int idx = 0; idx < listOfLeads.Count; idx++)
                        {
                            if (
                                (checkBox1.Checked == true) && (listOfLeads[idx].getLeadType() == LeadType.Email)
                                ||
                                (checkBox2.Checked == true) && (listOfLeads[idx].getLeadType() == LeadType.Phone)
                                ||
                                (checkBox3.Checked == true) && (listOfLeads[idx].getLeadType() == LeadType.Skype)
                                )
                            {
                                ListViewItem oItem = new ListViewItem();

                                oItem.Text = listOfLeads[idx].getLeadId().ToString();
                                oItem.SubItems.Add(listOfLeads[idx].getLeadTimeStamp().ToString(@"dd/MM/yyyy HH:mm:ss"));
                                oItem.SubItems.Add(listOfLeads[idx].getLeadType().ToString());
                                oItem.SubItems.Add(listOfLeads[idx].getLeadContent());
                                oItem.SubItems.Add(listOfLeads[idx].getLeadURL());
                                listView3.Items.Add(oItem);

                                nOfLeads++;
                            }

                        } // for

                        if (ckbAutoRefresh1.Checked == true)
                        {

                            // ensure visible
                            if (listView3.Items.Count - 1 > 0)
                            {
                                listView3.EnsureVisible(listView3.Items.Count - 1);
                            }
                        }

                    }
                    finally
                    {
                        listView3.EndUpdate();
                    }

                }

            }
           
            Int32 nNumberOfEmailsFound = 0;
            Int32 nNumberOfPhonesFound = 0;
            Int32 nNumberOfSkypesFound = 0;

            toolStripStatusLabel2.Text = listOfURLsTOBEProcessed.Count().ToString();

            toolStripStatusLabel5.Text = listOfURLsProcessed.Count().ToString();

            toolStripStatusLabel9.Text = listOfLeads.Count().ToString();

            for (Int32 idx = 0; idx < listOfLeads.Count(); idx++)
            {
                if (listOfLeads[idx].getLeadType() == LeadType.Email)
                    nNumberOfEmailsFound++;

                if (listOfLeads[idx].getLeadType() == LeadType.Phone)
                    nNumberOfPhonesFound++;

                if (listOfLeads[idx].getLeadType() == LeadType.Skype)
                    nNumberOfSkypesFound++;
            }

            toolStripStatusLabel12.Text = nNumberOfEmailsFound.ToString();

            toolStripStatusLabel15.Text = nNumberOfPhonesFound.ToString();

            toolStripStatusLabel18.Text = nNumberOfSkypesFound.ToString();

            IsRefreshLeadsList = true;

            if (nOfLeads > Convert.ToInt32(comboMaxLeadsToExtract.Text))
            {                    
                nState = -1;
                ButtonsState();
                MessageBox.Show("According to general settings, the max leads to extract is set to " + Convert.ToInt32(comboMaxLeadsToExtract.Text) + "." + Environment.NewLine + "They are extracted successfully. The process has stopped.", "LeadsExtractor");
            }                
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.linkLabel1.LinkVisited = true;
            
            System.Diagnostics.Process.Start(strNavigateLink);
        }

        private void openSelectedURLInBrowserToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Open selected URL in browser

            try
            {
                String strTheUrlToOpen = "";

                if (listView3.SelectedItems.Count == 1)
                {
                    strTheUrlToOpen = listView3.SelectedItems[0].SubItems[4].Text;
                    System.Diagnostics.Process.Start(strTheUrlToOpen);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            // remove item from lv_ExcludeEmailsWithName

            foreach (ListViewItem listViewItem in lv_ExcludeEmailsWithName.SelectedItems)
            {
                listViewItem.Remove();
                textBox1.Text = "";
            }
        }

        private void btnAdd1_Click(object sender, EventArgs e)
        {
            // add item to lv_ExcludeEmailsWithName

            if (textBox1.Text.Length < 3)
            {
                MessageBox.Show("You have to enter at least 3 characters string.","LeadsExtractor");
            }
            else
            {
                ListViewItem item = new ListViewItem(textBox1.Text);
                lv_ExcludeEmailsWithName.Items.Add(item);
                textBox1.Text = "";

            }

            // ensure visible
            if (lv_ExcludeEmailsWithName.Items.Count - 1 > 0)
            {
                lv_ExcludeEmailsWithName.EnsureVisible(lv_ExcludeEmailsWithName.Items.Count - 1);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            // edit item in lv_ExcludeEmailsWithName

            if (textBox1.Text.Length < 3)
            {
                MessageBox.Show("You have to enter at least 3 characters string.", "LeadsExtractor");
            }
            else
            {
                lv_ExcludeEmailsWithName.SelectedItems[0].Text = textBox1.Text;
                textBox1.Text = "";
            }            
        }

        private void lv_ExcludeEmailsWithName_SelectedIndexChanged(object sender, EventArgs e)
        {
            ListViewItem item = lv_ExcludeEmailsWithName.SelectedItems[0];
            textBox1.Text = item.SubItems[0].Text;
        }

        private void button10_Click(object sender, EventArgs e)
        {
            // add item to lv_ExcludeEmailsWithDomain

            if (textBox3.Text.Length < 3)
            {
                MessageBox.Show("You have to enter at least 3 characters string.", "LeadsExtractor");
            }
            else
            {
                ListViewItem item = new ListViewItem(textBox3.Text);
                lv_ExcludeEmailsWithDomain.Items.Add(item);
                textBox3.Text = "";
            }

            // ensure visible
            if (lv_ExcludeEmailsWithDomain.Items.Count - 1 > 0)
            {
                lv_ExcludeEmailsWithDomain.EnsureVisible(lv_ExcludeEmailsWithDomain.Items.Count - 1);
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            // add item to lv_ExcludeURLsWith

            if (textBox4.Text.Length < 3)
            {
                MessageBox.Show("You have to enter at least 3 characters string.", "LeadsExtractor");
            }
            else
            {
                ListViewItem item = new ListViewItem(textBox4.Text);
                lv_ExcludeURLsWith.Items.Add(item);
                textBox4.Text = "";
            }

            // ensure visible
            if (lv_ExcludeURLsWith.Items.Count - 1 > 0)
            {
                lv_ExcludeURLsWith.EnsureVisible(lv_ExcludeURLsWith.Items.Count - 1);
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            // edit item in lv_ExcludeEmailsWithDomain

            if (textBox3.Text.Length < 3)
            {
                MessageBox.Show("You have to enter at least 3 characters string.", "LeadsExtractor");
            }
            else
            {
                lv_ExcludeEmailsWithDomain.SelectedItems[0].Text = textBox3.Text;
                textBox3.Text = "";
            }     
        }

        private void button12_Click(object sender, EventArgs e)
        {
            // edit item in lv_ExcludeURLsWith

            if (textBox4.Text.Length < 3)
            {
                MessageBox.Show("You have to enter at least 3 characters string.", "LeadsExtractor");
            }
            else
            {
                lv_ExcludeURLsWith.SelectedItems[0].Text = textBox4.Text;
                textBox4.Text = "";
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            // remove item from lv_ExcludeEmailsWithDomain

            foreach (ListViewItem listViewItem in lv_ExcludeEmailsWithDomain.SelectedItems)
            {
                listViewItem.Remove();
                textBox3.Text = "";
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            // remove item from lv_ExcludeURLsWith

            foreach (ListViewItem listViewItem in lv_ExcludeURLsWith.SelectedItems)
            {
                listViewItem.Remove();
                textBox4.Text = "";
            }
        }

        private void lv_ExcludeEmailsWithDomain_SelectedIndexChanged(object sender, EventArgs e)
        {
            ListViewItem item = lv_ExcludeEmailsWithDomain.SelectedItems[0];
            textBox3.Text = item.SubItems[0].Text;
        }

        private void lv_ExcludeURLsWith_SelectedIndexChanged(object sender, EventArgs e)
        {
            ListViewItem item = lv_ExcludeURLsWith.SelectedItems[0];
            textBox4.Text = item.SubItems[0].Text;
        }

        private void button14_Click(object sender, EventArgs e)
        {
            // open autosave folder

            try
            {
                System.Diagnostics.Process.Start("explorer.exe", strAutosaveFolder);
            }
            catch (Exception ex)
            {
                //
            }

        }

        private void timer3_Tick(object sender, EventArgs e)
        {
            if (!Directory.Exists(strAutosaveFolder))
            {
                System.IO.Directory.CreateDirectory(strAutosaveFolder);
            }

            String strTimeStamp = string.Format("{0:yyyyMMdd_HHmmss_fff}", DateTime.Now);

            autoSaveLeads(strAutosaveFolder + "\\" + " leads_extractor_autosave_" + strTimeStamp + ".xlsx");
        }

        private void cbxAutoSaveResults_SelectedIndexChanged(object sender, EventArgs e)
        {
            timer3.Interval = 60 * 1000 * Convert.ToInt32(cbxAutoSaveResults.Text);           
        }

        private void timer4_Tick(object sender, EventArgs e)
        {
            // update statistics
            ++nMinutesOfRun;
            chart1.Series["Parsed URLs"].Points.AddXY(nMinutesOfRun, Convert.ToInt32(toolStripStatusLabel5.Text));
            chart2.Series["Leads"].Points.AddXY(nMinutesOfRun, Convert.ToInt32(toolStripStatusLabel9.Text));
            chart3.Series["Emails"].Points.AddXY(nMinutesOfRun, Convert.ToInt32(toolStripStatusLabel12.Text));
            chart3.Series["Phones"].Points.AddXY(nMinutesOfRun, Convert.ToInt32(toolStripStatusLabel15.Text));
            chart3.Series["Skypes"].Points.AddXY(nMinutesOfRun, Convert.ToInt32(toolStripStatusLabel18.Text));
            
        }

    }
}
