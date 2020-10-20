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
        internal static void AddQueries(string[] TableNamesArray, Excel.Workbook wbk) //Create queries for each table in TableNamesArray
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
                    if (ExistConnection(wbk, ConNamePrefix + TableName))
                    {
                        MessageBox.Show("Connection" + TableName + " already exists. Please delete existing connections");
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
    Dim result As Variant
    Dim CombineQueryName As String
    Dim CombineTableName As String
    Dim BareTableNames As String
    CombineQueryName = ""WBSCombineQuery""
    CombineTableName = ""TAB"" & CombineQueryName
   TableNames = Array({TableNames})
    BareTableNames = Join(TableNames, "","")

    For Each TableName In TableNames
        On Error Resume Next
        result = Empty
        result = ActiveWorkbook.Queries(TableName)
        On Error GoTo 0

        If IsEmpty(result) Then
            ActiveWorkbook.Queries.Add _
                Name:= TableName, _
                 Formula:= ""let Source = Excel.CurrentWorkbook(){{[Name="""""" & TableName & """"""]}}[Content] in Source""
        End If
    Next

    On Error Resume Next
    Application.DisplayAlerts = False
    ActiveWorkbook.Queries(CombineQueryName).Delete
    ActiveWorkbook.Worksheets(CombineQueryName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    DeleteList(CombineTableName)

    ActiveWorkbook.Queries.Add Name:= CombineQueryName, Formula:= _
        ""let"" & vbCrLf & ""    Source = Table.Combine({{"" & BareTableNames & ""}})"" _
        & vbCrLf & ""in"" & vbCrLf & ""    Source""
    ActiveWorkbook.Worksheets.Add
    ActiveSheet.Name = CombineQueryName
    With ActiveSheet.ListObjects.Add( _
        SourceType:= 0, _
        Source:= ""OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location="" & _
            CombineQueryName & "";Extended Properties="""""""""", _
        Destination:= Range(""$A$3"")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array(""SELECT * FROM ["" & CombineQueryName & ""]"")
        .RowNumbers = False
        .ListObject.DisplayName = CombineTableName
        .Refresh BackgroundQuery:= False
    End With

    End Sub

Sub DeleteList(ListName As String)
    Dim WS As Worksheet
    Dim result As Variant
    For Each WS In ActiveWorkbook.Worksheets
        On Error Resume Next
        WS.ListObjects(ListName).Delete
        On Error GoTo 0
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
        internal static void CombineWBS()
        {
            try
            {
                Excel.Application xlAp = Globals.ThisAddIn.Application;
                Excel.Workbook XlWb = xlAp.ActiveWorkbook;
                Excel.Worksheet WBSCombinedSht;
                Excel.QueryTable QT;
                List<string> WBSTables = new List<string>();

                //VBA project object model needs to be trusted for this to work
                if (VBATrusted(XlWb) == false) 
                {
                    MessageBox.Show("No Access to VB Project\n\rPlease allow access in Trusted Sources\n\r" +
                        "File > Options > Trust Center > Trust Center Settings > Macro Settings > Trust Access to the VBA project object model");
                    return;
                }

                //Collect all WBS table names and create queries
                xlAp.ScreenUpdating = false;
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
                AddQueries(WBSTables.ToArray(), XlWb);

                //todo: Create append query to combined all WBS queries into one
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

                //Collect all WBS table names
                xlAp.ScreenUpdating = true;
            }
            catch (Exception ex)
            {
                Globals.ThisAddIn.ExMsg(ex);
            }
        }
        internal static bool VBATrusted(Excel.Workbook xlWb) //Check if VBA project object model is trusted
        {
            try
            {
                string VBProjName = xlWb.VBProject.Name;
                return true;
            }
            catch (Exception ex)
            {
                if ((uint)ex.HResult == 0x800a03ec)
                {
                    return false;
                }
                else
                {
                    Globals.ThisAddIn.ExMsg(ex);
                    return false;
                }
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
        internal static bool ExistConnection(Excel.Workbook XlWb, string QueryName)
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
