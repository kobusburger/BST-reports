using Microsoft.Vbe.Interop;
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
        internal static string BSTPath = "%USERPROFILE%\\AppData\\Local\\BSTEnterprise\\InquiryReports\\Proddbhttpzadc1pbst02.zutari.com";

        internal static void ParseWBS(BSTMonitorForm BSTMonForm)
        {
            try
            {
                Excel.Application xlAp = Globals.ThisAddIn.Application;
                Excel.Workbook XlWb = xlAp.ActiveWorkbook;
                Excel.Worksheet XlSh;
                Excel.QueryTable QT;
                long CurrentRow = 0; long LastRow = 0;
                string ProjNo = ""; string ProjName = "";
                string[] Project; string ProjCellText;
                string ConnectionString;
                string WBSTableName;

 //Import BST report
                xlAp.ScreenUpdating = false;
                XlSh = XlWb.Sheets.Add();
                ConnectionString = "FINDER;file:///" + Environment.ExpandEnvironmentVariables(BSTPath + "\\PrjWbs.htm");
                QT = XlSh.QueryTables.Add(Connection: ConnectionString, Destination: XlSh.Range["$A$1"]);
                QT.Refresh(false);
                QT.Delete();

//Inititalise BST variables
                LastRow = XlSh.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                ProjCellText = XlSh.Range["B16"].Text;
                char[] delimeter = { '-' }; //I do not know why I had to create a delimeter variable instead of usind {'-'} as the first argument
                Project = ProjCellText.Split(delimeter , 2);
                if (Project.Length > 0)
                {
                    ProjNo = Project[0].Trim();
                    ProjName = Project[1].Trim();
                }

//allocate table and sheet names
                WBSTableName = "WBS" + ProjNo;
                int Counter = 0;
                string PrevWBSTableName = WBSTableName;
                while (ExistListObject(XlWb, "Tab"+WBSTableName)) //Check if the table name exists
                {
                    Counter += 1;
                    WBSTableName = PrevWBSTableName + "_" + Counter;
                }
                XlSh.Name = WBSTableName;

//Parse report
                LastRow -= DeleteRows(XlSh, 20, 21); //Delete rows between headings and first table rows
                LastRow -= DeleteRows(XlSh, 2, 18); //Delete rows between report name and headings
                XlSh.Range["A:B"].Insert(); // Insert 2 columns
                XlSh.Range["2:2"].Insert(); //Insert blank row above headers
                XlSh.Cells[3,1].Value = "Project";
                XlSh.Cells[3,2].Value = "name";

                CurrentRow = 4; //First table row
                while (CurrentRow <= LastRow)
                {
                    if (LastRow % 10 == 0)
                        xlAp.StatusBar = string.Format("Progress: {0:f0}%", CurrentRow * 100 / LastRow);

                    // Identify row type
                    if (XlSh.Cells[CurrentRow, 1].Text.Trim() != "")
                    {
                        XlSh.Cells[CurrentRow, 1].Value = ProjNo;
                        XlSh.Cells[CurrentRow, 2].Value = ProjName;
                    }
                    else if (XlSh.Cells[CurrentRow, 3].Text.Trim() == "" && XlSh.Cells[CurrentRow + 1, 3].Text.Trim() == "Project WBS Report") //Page break
                    {
                        DeleteRows(XlSh, CurrentRow, CurrentRow + 8); //Delete rows between pages
                    }
                    else if (XlSh.Cells[CurrentRow, 3].Text.Trim() == "" && XlSh.Cells[CurrentRow + 2, 3].Text.Trim() == "Totals") //Report end
                    {
                        DeleteRows(XlSh, CurrentRow, CurrentRow + 2); //Delete rows at the end
                    break;
                    }
                    CurrentRow += 1;
                }
                XlSh.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, XlSh.Cells[3, 1].CurrentRegion,false, 
                    Excel.XlYesNoGuess.xlYes).name ="Tab" + WBSTableName;
                xlAp.StatusBar = false;
                xlAp.ScreenUpdating = true;
                BSTMonForm.FileEvents.AppendText(WBSTableName + " added: " + "\r\n");
            }
            catch (Exception ex)
            {
                Globals.ThisAddIn.ExMsg(ex);
            }
        }
        internal static void ParseAnalysis(BSTMonitorForm BSTMonForm)
        {
            try
            {
                Excel.Application xlAp = Globals.ThisAddIn.Application;
                Excel.Workbook XlWb = xlAp.ActiveWorkbook;
                Excel.Worksheet XlSh;
                Excel.QueryTable QT;
                long CurrentRow = 0; long LastRow = 0;
                string ConnectionString;
                string AnalTableName;

                //Import BST report
                xlAp.ScreenUpdating = false;
                XlSh = XlWb.Sheets.Add();
                ConnectionString = "FINDER;file:///" + Environment.ExpandEnvironmentVariables(BSTPath + "\\PrjAnalysis.htm");
                QT = XlSh.QueryTables.Add(Connection: ConnectionString, Destination: XlSh.Range["$A$1"]);
                QT.Refresh(false);
                QT.Delete();

//                return;
                //Inititalise BST variables
                LastRow = XlSh.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

                //allocate table and sheet names
                AnalTableName = "Analysis";
                if (ExistSheet(XlWb, AnalTableName))
                {
                    xlAp.DisplayAlerts = false;
                    XlWb.Sheets[AnalTableName].Delete();
                    xlAp.DisplayAlerts = true;
                }
                XlSh.Name = AnalTableName;

                //Parse report
                LastRow -= DeleteRows(XlSh, 18, 19); //Delete rows between headings and first table rows
                LastRow -= DeleteRows(XlSh, 2, 16); //Delete rows between report name and headings
                XlSh.Range["2:2"].Insert(); //Insert blank row above headers

                CurrentRow = 4; //First table row
                while (CurrentRow <= LastRow)
                {
                    if (LastRow % 10 == 0)
                        xlAp.StatusBar = string.Format("Progress: {0:f0}%", CurrentRow * 100 / LastRow);

                    // Identify row type
                    if (XlSh.Cells[CurrentRow, 1].Text.Trim() != "")
                    { }
                    else if (XlSh.Cells[CurrentRow, 1].Text.Trim() == "" && XlSh.Cells[CurrentRow + 1, 1].Text.Trim() == "Project Analysis Report") //Page break
                    {
                        DeleteRows(XlSh, CurrentRow, CurrentRow + 8); //Delete rows between pages
                    }
                    else if (XlSh.Cells[CurrentRow, 1].Text.Trim() == "" && XlSh.Cells[CurrentRow + 2, 1].Text.Trim() == "Totals") //Report end
                    {
                        DeleteRows(XlSh, CurrentRow, CurrentRow + 4); //Delete rows at the end
                        break;
                    }
                    CurrentRow += 1;
                }
                XlSh.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, XlSh.Cells[3, 1].CurrentRegion, false,
                    Excel.XlYesNoGuess.xlYes).name = "Tab" + AnalTableName;
                xlAp.StatusBar = false;
                xlAp.ScreenUpdating = true;
                BSTMonForm.FileEvents.AppendText(AnalTableName + " added: " + "\r\n");
            }
            catch (Exception ex)
            {
                Globals.ThisAddIn.ExMsg(ex);
            }
        }
        internal static void CombineWBS()
            //todo: test addquery
        {
            Excel.Application xlAp = Globals.ThisAddIn.Application;
            Excel.Workbook XlWb = xlAp.ActiveWorkbook;
            string[] TableNames = { "WBSTable1", "abcd1" };
            AddQuery(TableNames, XlWb);

        }
        internal static void AddQuery(string[] TableNamesArray, Excel.Workbook wbk) //Create queries for each table in TableNamesArray
        //https://stackoverflow.com/questions/61622872/adding-power-queries-to-excel-using-c-sharp
        //https://docs.microsoft.com/en-us/office/vba/language/reference/visual-basic-add-in-model/objects-visual-basic-add-in-model#vbcomponent
        //https://stackoverflow.com/questions/64210190/how-to-create-queries-and-connections#
        //http://www.cpearson.com/excel/vbe.aspx
        {
            try
            {
                string MacroName ;
                string wbkName = wbk.Name;
                string ConNamePrefix = "Query - ";
                string TableNames;
                VBComponent newStandardModule;
                string VBAcodeText;

                foreach (string TableName in TableNamesArray) //Return if query already exists
                {
                    if (ExistQuery(wbk, ConNamePrefix + TableName))
                    {
                        MessageBox.Show(ConNamePrefix + TableName + " already exists");
                        return;
                    }
                }

                Random RandNo = new Random();
                MacroName = "Addquery" + RandNo.Next(100000, 1000000); //Randomize the macro name
                TableNames = "\"" + string.Join("\",\"", TableNamesArray) + "\""; //Create string in VBA expected format
                newStandardModule = wbk.VBProject.VBComponents.Add(Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_StdModule);

                var codeModule = newStandardModule.CodeModule;

                // add vba code to module
                VBAcodeText = $@"
Sub {MacroName}()
    Dim TableName As Variant
    Dim TableNames() As Variant
    TableNames = Array({TableNames})

    For Each TableName In TableNames
        ActiveWorkbook.Queries.Add _
            Name:= TableName, _
            Formula:= ""let Source = Excel.CurrentWorkbook(){{[Name="""""" & TableName & """"""]}}[Content] in Source""

            Workbooks(""{wbkName}"").Connections.Add2 _
            Name:= ""{ConNamePrefix}"" & TableName, _
            Description:= ""Connection to the "" & TableName & "" query in the workbook."", _
            ConnectionString:= ""OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location="" & TableName & "";Extended Properties="", _
            CommandText:= ""SELECT * FROM ["""" & TableName & """"]"", _
            lCmdtype:= 2
    Next
End Sub
                ";
                codeModule.InsertLines(4, VBAcodeText);
                wbk.Application.Run($@"{newStandardModule.Name}.{MacroName}");

                wbk.VBProject.VBComponents.Remove(newStandardModule);
            }
            catch (Exception ex)
            {
                Globals.ThisAddIn.ExMsg(ex);
            }
        }

        /*        internal static void CombineWBS()
                {
                    try
                    {
                        Excel.Application xlAp = Globals.ThisAddIn.Application;
                        Excel.Workbook XlWb = xlAp.ActiveWorkbook;
                        Excel.Worksheet XlSh;
                        Excel.QueryTable QT;
                        long CurrentRow = 0; long LastRow = 0;
                        string ProjNo = ""; string ProjName = "";
                        string[] Project; string ProjCellText;
                        string ConnectionString;
                        string WBSShtName;
                        List<string> WBSTables = new List<string>();

                        xlAp.ScreenUpdating = false;
                        WBSShtName = "WBS Combined";
                        if (ExistSheet(XlWb, WBSShtName))
                        {
                            xlAp.DisplayAlerts = false;
                            XlWb.Sheets[WBSShtName].Delete();
                            xlAp.DisplayAlerts = true;
                        }
                        XlSh = XlWb.Sheets.Add();
                        XlSh.Name = WBSShtName;

                            int NoShts = XlWb.Worksheets.Count;
                        foreach (Excel.Worksheet Sheet in XlWb.Worksheets)
                        {
                            int NoQ = Sheet.QueryTables.Count;
                            foreach (Excel.QueryTable Table in Sheet.QueryTables)
                            {
                                string QName = Table.Name;
                                    WBSTables.Add(Table.Name);
                            }
                        }

                        int NoCon = XlWb.Connections.Count;
                        foreach (Excel.WorkbookConnection Con in XlWb.Connections)
                        {
                            string ConName = Con.Name;
                            string Descr = Con.Description;
                            bool inmodel = Con.InModel;
                            string connection = Con.OLEDBConnection.Connection;
                            string file = Con.OLEDBConnection.SourceDataFile;
                            int range = Con.Ranges.Count;


                            WBSTables.Add(Con.Name);
                        }

                        //Collect all WBS table names
                        foreach (Excel.Worksheet Sheet in XlWb.Worksheets)
                        {
                            int NoLists = Sheet.ListObjects.Count;
                            foreach (Excel.ListObject Table in Sheet.ListObjects)
                            {
                                string ListName = Table.Name;
                                if (Table.Name.Substring(0, 6) == "TabWBS")
                                {
                                    WBSTables.Add(Table.Name); //Add table to combine
                                }
                            }
                        }
                    xlAp.ScreenUpdating = true;
                    }
                    catch (Exception ex)
                    {
                        Globals.ThisAddIn.ExMsg(ex);
                    }
                }*/

        //From https://stackoverflow.com/questions/61622872/adding-power-queries-to-excel-using-c-sharp
        /*        public void AddQuery(string m_script_path, string query_name, Excel.Workbook wk)
                {
                    VBComponent newStandardModule;
                    if (wk.VBProject.VBComponents.Count == 0)
                    {
                        newStandardModule = wk.VBProject.VBComponents.Add(Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_StdModule);
                    }
                    else
                    {
                        newStandardModule = wk.VBProject.VBComponents.Item(1);
                    }

                    var codeModule = newStandardModule.CodeModule;

                    // add vba code to module
                    var lineNum = codeModule.CountOfLines + 1;
                    var macroName = "addQuery";
                    var codeText = "Public Sub " + macroName + "()" + "\r\n";
                    codeText += "M_Script = CreateObject(\"Scripting.FileSystemObject\").OpenTextFile(\"" + m_script_path + "\", 1).ReadAll" + "\r\n";
                    codeText += "ActiveWorkbook.Queries.Add Name:=\"" + query_name + "\", Formula:=M_Script\r\n";
                    codeText += "ActiveWorkbook.Connections.Add2 _\r\n";
                    codeText += "\"Query - test\", _\r\n";
                    codeText += "\"Connection to the '" + query_name + "' query in the workbook.\", _\r\n";
                    codeText += "\"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" + query_name + ";Extended Properties=\" _\r\n";
                    codeText += ", \"\"\"" + query_name + "\"\"\", 6, True, False\r\n";

                    codeText += "End Sub";

                    codeModule.InsertLines(lineNum, codeText);

                    var macro = string.Format("{0}.{1}", newStandardModule.Name, macroName);

                    wk.Application.Run(macro);

                    codeModule.DeleteLines(lineNum, 9);
                }*/
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
        internal static bool ExistListObject(Excel.Workbook XlWb, string ListName)
        {
            // Returns true if a list object exist in the workbook
            foreach (Excel.Worksheet Sheet in XlWb.Worksheets) // Loop through all the worksheets
            {
                foreach (Excel.ListObject ListObj in Sheet.ListObjects) // Loop through each table in the worksheet
                {
                    if (ListObj.Name == ListName)
                    {
                        return true;
                    }
                }
            }

            return false;
        }
        internal static bool ExistSheet(Excel.Workbook XlWb, string SheetName)
        {
            // Returns true if a sheet exists in the workbook
            foreach (Excel.Worksheet Sheet in XlWb.Worksheets) // Loop through all the worksheets
            {
                if (Sheet.Name == SheetName)
                {
                    return true;
                }
            }
            return false;
        }
        internal static bool ExistQuery(Excel.Workbook XlWb, string QueryName)
        {
            // Returns true if a query exists in the workbook
            foreach (Excel.WorkbookConnection Query in XlWb.Connections) // Loop through all the worksheets
            {
                if (Query.Name == QueryName)
                {
                    return true;
                }
            }
            return false;
        }
    }
}
