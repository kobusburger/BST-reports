using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace BST_reports
{
    static class BST
    {
        static string BSTTableName = "BSTData";
        static string BSTPath = "%USERPROFILE%\\AppData\\Local\\BSTEnterprise\\InquiryReports\\Proddbhttpzadc1pbst02.zutari.com";

        internal static void ArrangeBSTCosts()
        {//Convert "Project Detail Charges" BST report into a data table
            try
            {
                long Ry=0; long LaasteRy=0;
                string Project=""; string ProjDescr="";
                string Phase=""; string PhaseDescr="";
                string Task=""; string TaskDescr="";
                string CostType=""; string EVCColCellText="";
                long EVCCol=0; long DetCol=0;
                Excel.Worksheet XlSh;
                Excel.Workbook XlWb;
                Excel.Application xlAp;
                string[] COlHdrs= { "Project", "Project Description", "Phase", "Phase Description", "Task", "Task Description", "Cost Type", "Description", "EVC Code", "Name", "Class / GL Acct", "Co", "Org", "Actv/ Unit", "Bill Ind", "Document Number", "Detail Type", "Transaction Date", "Period End Date", "Reg / OT", "Hours / Quantity", "Cost rate", "Cost Amount", "Effort Rate", "Effort Amount" };
                int ColNo;
                string TempStr;

//Initialising
                xlAp = Globals.ThisAddIn.Application;
                XlWb = xlAp.ActiveWorkbook;
                XlSh = XlWb.ActiveSheet;
                xlAp.StatusBar = "Progress: Initialising";
                Globals.ThisAddIn.LogTrackInfo("ArrangeBSTCosts");

                if (Pivot.ExistListObject(BSTTableName)) //Check if the table name exists
                {
                    MessageBox.Show(BSTTableName + " table already exist");
                    return; // exit if the report is alreayd converted to a data table
                }

                if (XlSh.Cells[3, 2].text != "Project Detail Charges")
                {
                    MessageBox.Show("This is not a 'Project Detail Charges' report\r\nThe report must be created via Project/ Reporting/ Project Detail Charges");
                    return; // exit if this is not a Project Detail Charges report
                }

                XlSh.Name = "BST";
                xlAp.ScreenUpdating = false;
                LaasteRy = XlSh.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                // Insert/ delete columns ------------------------------------------------------------------------------
                for (ColNo = 1; ColNo <= 8; ColNo++)    // Insert columns up to "EVC Code" column
                    XlSh.Cells[1, 1].EntireColumn.Insert();
                XlSh.Columns[4 + 8].delete();  // Delete "Task" column because a new one is added

// Assign headings --------------------------------------------------------------------------------------
                ColNo = 1;
                foreach (var Hdr in COlHdrs)
                {
                    XlSh.Cells[1, ColNo] = Hdr;
                    if (Hdr == "EVC Code")
                        EVCCol = ColNo;
                    if (Hdr == "Detail Type")
                        DetCol = ColNo;
                    ColNo += 1;
                }

// Process each row -------------------------------------------------------------------------------------
                Ry = 2;
                while (Ry <= LaasteRy)
                {
                    if (LaasteRy % 100 == 0)
                        xlAp.StatusBar = string.Format("Progress: {0:f0}%", Ry * 100 / LaasteRy);
                    // Identify row type
                    EVCColCellText = XlSh.Cells[Ry, EVCCol].text;
                    if (EVCColCellText.Length > 9 && EVCColCellText.Substring(0, 9) == "Project :") //The second expression is not evaluated if the first is false (short circuit evaluation)
                        {
                            TempStr = EVCColCellText;
                            Project = TempStr.Substring(13, TempStr.IndexOf("-") - 13).Trim();
                            ProjDescr = TempStr.Substring(TempStr.IndexOf("-") - 2).Trim();
                            XlSh.Rows[Ry].EntireRow.Delete();
                            LaasteRy -= 1;
                            Ry -= 1;
                        }
                    else if (EVCColCellText.Length > 7 && EVCColCellText.Substring(0, 7) == "Phase :")
                        {
                            TempStr = EVCColCellText;
                            Phase = TempStr.Substring(11, TempStr.IndexOf("-") - 11).Trim();
                            PhaseDescr = TempStr.Substring(TempStr.IndexOf("-") - 2).Trim();
                            XlSh.Rows[Ry].EntireRow.Delete();
                            LaasteRy -= 1;
                            Ry -= 1;
                        }
                    else if (EVCColCellText.Length > 6 && EVCColCellText.Substring(0, 6) == "Task :")
                        {
                            TempStr = EVCColCellText;
                            Task = TempStr.Substring(10, TempStr.IndexOf("-") - 10).Trim();
                            TaskDescr = TempStr.Substring(TempStr.IndexOf("-") - 2).Trim();
                            XlSh.Rows[Ry].EntireRow.Delete();
                            LaasteRy -= 1;
                            Ry -= 1;
                        }
                    else if (EVCColCellText == "Labor" | EVCColCellText == "Expense")
                    {
                        CostType = EVCColCellText;
                        XlSh.Rows[Ry].EntireRow.Delete();
                        LaasteRy -= 1;
                        Ry -= 1;
                    }
                    else if (Array.IndexOf(new[] { "P", "E", "R", "U", "M" }, (XlSh.Cells[Ry, DetCol].text)) >= 0)
                    {
                        XlSh.Cells[Ry, 1] = Project;
                        XlSh.Cells[Ry, 2] = ProjDescr;
                        XlSh.Cells[Ry, 3] = Phase;
                        XlSh.Cells[Ry, 4] = PhaseDescr;
                        XlSh.Cells[Ry, 5] = Task;
                        XlSh.Cells[Ry, 6] = TaskDescr;
                        XlSh.Cells[Ry, 7] = CostType;
                        if (Array.IndexOf(new[] { "P", "E", "U" }, (XlSh.Cells[Ry, DetCol].text)) >= 0)
                            XlSh.Cells[Ry, DetCol + 3].Insert(Excel.XlInsertShiftDirection.xlShiftToRight);

                        // Move Description into the item row
                        if (!double.TryParse(XlSh.Cells[Ry + 1, EVCCol].text, out double number) && string.IsNullOrEmpty(XlSh.Cells[Ry + 1, EVCCol + 1].text))
                        {
                            XlSh.Cells[Ry, EVCCol - 1] = XlSh.Cells[Ry+1, EVCCol].text;
                            XlSh.Rows[Ry + 1].EntireRow.Delete();
                            LaasteRy -= 1;
                        }
                    }
                    else // delete row
                    {
                        XlSh.Rows[Ry].EntireRow.Delete();
                        LaasteRy -= 1;
                        Ry -= 1;
                    }

                    Ry += 1;
                }

// Create BST table -------------------------------------------------------------------------------
                XlSh.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, XlSh.Range["A1"].CurrentRegion, default, Excel.XlYesNoGuess.xlYes).Name = BSTTableName;

                // Add month column formula
                XlSh.Range["Z1"].Value = "Month";
                XlSh.Range["Z2"].Value = "=TEXT(R2,\"yyyy-mm\")";
                XlSh.Range["Z2", "Z" + LaasteRy].FillDown();
                XlSh.Range["R2", "S" + LaasteRy].NumberFormat = "yyyy-mm-dd"; // Format dates

                xlAp.StatusBar = false;
                xlAp.ScreenUpdating = true;
            }

            catch (Exception ex)
            {
                Globals.ThisAddIn.ExMsg(ex);
            }
        }
        internal static void ParseWBS()
        {
            try
            {
                Excel.Application xlAp = Globals.ThisAddIn.Application;
                Excel.Workbook XlWb = xlAp.ActiveWorkbook;
                Excel.Worksheet XlSh;
                Excel.QueryTable QT;
                long CurrentRow = 0; long LastRow = 0;
                string ProjNo = ""; string ProjName = "";
                string Project;
                string ConnectionString;

//                xlAp.ScreenUpdating = false;
                XlSh = XlWb.Sheets.Add();
                ConnectionString = "FINDER;file:///" + Environment.ExpandEnvironmentVariables(BSTPath + "\\PrjWbs.htm");
                QT = XlSh.QueryTables.Add(Connection: ConnectionString, Destination: XlSh.Range["$A$1"]);
                QT.Refresh(false);
                QT.Delete();

                LastRow = XlSh.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                Project = XlSh.Range["B25"].Text;
                ProjNo = Project.Split('-')[0];
                ProjName = Project.Split('-')[1];

                XlSh.Range["A:B"].Insert();// Insert 2 columns
                LastRow -= DeleteRows(XlSh, 28, 29); //Delete rows between headings and first table rows
                LastRow -= DeleteRows(XlSh, 4, 26); //Delete rows between report name and headings

                CurrentRow = 5; //First table row
                while (CurrentRow <= LastRow)
                {
                    if (LastRow % 100 == 0)
                        xlAp.StatusBar = string.Format("Progress: {0:f0}%", CurrentRow * 100 / LastRow);

                    // Identify row type
                    if (XlSh.Cells[CurrentRow, 3] == "" && XlSh.Cells[CurrentRow + 1, 3] == "") //Page break
                    {
                        XlSh.Cells[CurrentRow, CurrentRow + 9].EntireRow.Delete(); //Delete rows between pages
                    }
                    else if (XlSh.Cells[CurrentRow, 3] == "" && XlSh.Cells[CurrentRow + 1, 3] == "Totals") //Report end
                    {
                        XlSh.Cells[CurrentRow, CurrentRow + 3].EntireRow.Delete(); //Delete rows between pages
                    }
                }
                xlAp.StatusBar = false;
                xlAp.ScreenUpdating = true;
            }
            catch (Exception ex)
            {
                Globals.ThisAddIn.ExMsg(ex);
            }
        }
        internal static void ParseAnalysis()
        {
            try
            {
                Excel.Application xlAp = Globals.ThisAddIn.Application;
                Excel.Workbook XlWb = xlAp.ActiveWorkbook;
                Excel.Worksheet XlSh;
                Excel.QueryTable QT;
                long CurrentRow = 0; long LastRow = 0;
                string ProjNo = ""; string ProjName = "";
                string Project;
                string ConnectionString;

                XlSh = XlWb.Sheets.Add();
                ConnectionString = "FINDER;file:///" + Environment.ExpandEnvironmentVariables(BSTPath + "\\PrjAnalysis.htm");
                QT = XlSh.QueryTables.Add(Connection: ConnectionString, Destination: XlSh.Range["$A$1"]);
                QT.Refresh(false);
                QT.Delete();
            }
            catch (Exception ex)
            {
                Globals.ThisAddIn.ExMsg(ex);
            }
        }
        internal static long DeleteRows(Excel.Worksheet XlSh, long StartRow, long EndRow)
        {
            long NoOfRows = 0;
            if (EndRow >= StartRow)
            {
                string DeleteRange;
                DeleteRange = StartRow + ":" + EndRow;
                XlSh.Range[DeleteRange].Delete();
                NoOfRows = EndRow - StartRow + 1;
            }
            return NoOfRows;
        }
        internal static void CombineWBS()
        { }
    }
}
