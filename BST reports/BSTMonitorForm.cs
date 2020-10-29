using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace BST_reports
{
    public partial class BSTMonitorForm : Form
    {
        System.DateTime ImportTime = DateTime.Now; //Used to skip events
        public static int FileEventCounter=0;
        public BSTMonitorForm()
        {
            InitializeComponent();
        }

        private void BSTMonitorForm_Load(object sender, EventArgs e)
        {
            try
            {
                this.BSTFileWatcher.Path = Environment.ExpandEnvironmentVariables(BST.BSTPath);
            }
            catch (Exception ex)
            {
                Globals.ThisAddIn.ExMsg(ex);
            }
        }

        private void BSTFileWatcher_Changed(object sender, System.IO.FileSystemEventArgs e)
        {
            ProcessReport(e.Name);
        }

        private void BSTFileWatcher_Created(object sender, System.IO.FileSystemEventArgs e)
        {
            ProcessReport(e.Name);
        }
        private void ProcessReport(string FilePath)
        {
            Excel.Worksheet whst;
            try
            {
                switch (Path.GetFileName(FilePath))
                {
                    case "PrjWbs.htm":
                       if ((DateTime.Now - ImportTime).TotalMilliseconds<2000) //There are two events. Ignore first event
                       {

                            whst = BST.ImportReport(FilePath);
                            string ProjNo = BST.ParseWBS(whst);
                            if (ProjNo != "")
                            {
                                this.FileEvents.AppendText("WBS" + ProjNo + " report added" + "\r\n");
                            }
                            FileEventCounter += 1;
                            EventCounter.Text = FileEventCounter.ToString();
                       }
                        ImportTime = DateTime.Now;
                        break;
                    case "PrjAnalysis.htm":
                        if ((DateTime.Now - ImportTime).TotalMilliseconds < 2000) //There are two events. Ignore first event
                        {
                            whst = BST.ImportReport(FilePath);
                            BST.ParseAnalysis(whst);
                            this.FileEvents.AppendText("Analysis report added" + "\r\n");
                            FileEventCounter += 1;
                            EventCounter.Text = FileEventCounter.ToString();
                        }
                        break;
                }
           }
            catch (Exception ex)
            {
                Globals.ThisAddIn.ExMsg(ex);
            }
        }
        private void CloseButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void BSTMonitorForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            FileEventCounter = 0;
            this.BSTFileWatcher.EnableRaisingEvents = false; //Stop file monitoring
        }
    }
}
