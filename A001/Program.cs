using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Diagnostics;

namespace LeadsExtractor
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            string strProcessName = Process.GetCurrentProcess().ProcessName;
            Process[] Oprocesses = Process.GetProcessesByName(strProcessName);

            if (Oprocesses.Length > 1)
            {
                MessageBox.Show("The application is already running.");
            }
            else
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new Form1());
            }
        }
    }
}
