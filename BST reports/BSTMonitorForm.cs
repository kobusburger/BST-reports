using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace BST_reports
{
    public partial class BSTMonitorForm : Form
    {
        int FileEventCounter=0;
        public BSTMonitorForm()
        {
            InitializeComponent();
        }

        private void BSTMonitorForm_Load(object sender, EventArgs e)
        {
        }

        private void BSTFileWatcher_Changed(object sender, System.IO.FileSystemEventArgs e)
        {
            ImportReport(e.Name);
        }

        private void BSTFileWatcher_Created(object sender, System.IO.FileSystemEventArgs e)
        {
            ImportReport(e.Name);
        }
        private void ImportReport(string FileName)
        {
            try
            {
                switch (FileName)
                {
                    case "PrjWbs.htm":
                        BST.ParseWBS();

                        FileEventCounter += 1;
                        EventCounter.Text = FileEventCounter.ToString();
                        FileEvents.AppendText(FileEventCounter + ": WBS report added: " + "\r\n");
                        break;
                    case "PrjAnalysis.htm":
                        BST.ParseAnalysis();

                        FileEventCounter += 1;
                        EventCounter.Text = FileEventCounter.ToString();
                        FileEvents.AppendText(FileEventCounter + ": Analysis report added: " + "\r\n");
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
    }
}
