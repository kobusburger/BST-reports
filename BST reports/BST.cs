using Microsoft.Office.Interop.Excel;
using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Excel = Microsoft.Office.Interop.Excel;

namespace BST_reports
{
    static class BST
    {
        internal static string BSTPath = "%USERPROFILE%\\AppData\\Local\\BSTEnterprise\\InquiryReports"; //\\Proddbhttpzadc1pbst02.zutari.com";
        internal static Excel.Worksheet ImportReport(string FileName) //Import the report onto a new sheet
        {
            Excel.Application xlAp = Globals.ThisAddIn.Application;
            Excel.Workbook XlWb = xlAp.ActiveWorkbook;
            Excel.Worksheet XlSh;
            Excel.QueryTable QT;
            string ConnectionString;

            try
            {
                //Import BST report
                xlAp.ScreenUpdating = false;
                XlSh = XlWb.Sheets.Add(After:XlWb.Worksheets[XlWb.Worksheets.Count]);
//                SetShtName(XlSh, "ImportDate", DateTime.Now.ToString("yyyy-MM-dd"));
                ConnectionString = "FINDER;file:///" + Environment.ExpandEnvironmentVariables(BSTPath + "\\" + FileName);
                QT = XlSh.QueryTables.Add(Connection: ConnectionString, Destination: XlSh.Range["$A$1"]);
                QT.WebSelectionType = Excel.XlWebSelectionType.xlEntirePage;
                QT.WebDisableDateRecognition = true;
                QT.Refresh(false);
                QT.Delete();
                xlAp.ScreenUpdating = true;
                return XlSh;
            }
            catch (Exception ex)
            {
                Globals.ThisAddIn.ExMsg(ex);
                return null;
            }
        }
        internal static string ParsePjWBS(Excel.Worksheet XlSh) //return the project number or blank if an error occured
        {
            try
            {
                Excel.Application xlAp = Globals.ThisAddIn.Application;
                Excel.Workbook XlWb = xlAp.ActiveWorkbook;
                long CurrentRow; long LastRow;
                long HeadingsRow;
                string ProjNo = ""; string ProjName = "";
                string[] Project; string ProjCellText;
                string WBSTableName;

//Inititalise BST variables
                ProjCellText = XlSh.Range["B25"].Text;
                char[] delimeter = { '-' }; //I do not know why I had to create a delimeter variable instead of usind {'-'} as the first argument
                Project = ProjCellText.Split(delimeter, 2);
                if (Project.Length > 0)
                {
                    ProjNo = Project[0].Trim();
                    ProjName = Project[1].Trim();
                }

                //allocate table and sheet names
                xlAp.ScreenUpdating = false;
                WBSTableName = "WBS" + ProjNo;
                if (ExistSheet(XlWb,WBSTableName))
                {
                    xlAp.DisplayAlerts = false;
                    XlWb.Worksheets[WBSTableName].delete();
                    xlAp.DisplayAlerts = true;
                }
                XlSh.Name = WBSTableName;

                //Parse report
                HeadingsRow = 27;
                LastRow = XlSh.UsedRange.Rows.Count;
                LastRow -= DeleteRows(XlSh, HeadingsRow+1, HeadingsRow + 2); //Delete rows between headings and first table rows
                //LastRow -= DeleteRows(XlSh, 4, HeadingsRow - 1); //Delete rows between report name and headings
                //HeadingsRow = 5;

                XlSh.Range["B5"].TextToColumns(    //Split date
                    Type.Missing, //Destination
                    XlTextParsingType.xlDelimited, //DataType
                    XlTextQualifier.xlTextQualifierDoubleQuote,    //TextQualifier
                    false,         // Consecutive Delimiter
                    Type.Missing,  // Tab
                    Type.Missing,  // Semicolon
                    false,         // Comma
                    false,         // Space
                    true,          // Other
                    ":",           // Other Char
                    Type.Missing,  // Field Info
                    Type.Missing,  // Decimal Separator
                    Type.Missing,  // Thousands Separator
                    Type.Missing); // Trailing Minus Numbers
                SetNameRange(XlWb, "ImportDateWBS"+WBSTableName, XlSh.Range["C5"]);
                XlSh.Range["C5"].Value=XlSh.Range["C5"].Text.Trim();

                XlSh.Range["A:B"].Insert(); // Insert 2 columns
                XlSh.Range[(HeadingsRow - 1) + ":" + (HeadingsRow - 1)].Clear();
                XlSh.Cells[HeadingsRow, 1].Value = "Project";
                XlSh.Cells[HeadingsRow, 2].Value = "Name";

                CurrentRow = HeadingsRow + 1; //First table row
                while (CurrentRow <= LastRow+1)
                {
                    if (CurrentRow % 10 == 0)
                        xlAp.StatusBar = string.Format("Progress: {0:f0}%", CurrentRow * 100 / LastRow);

                    // Identify row type
                    string SwitchText = XlSh.Cells[CurrentRow, 3].Text.Trim();
                    switch (SwitchText)
                    {
                        case "Project WBS Report": //Delete page breaks
                            LastRow -= DeleteRows(XlSh, CurrentRow - 3, CurrentRow + 6);
                            CurrentRow -= 4; //3+1 because it is incremented by 1 later
                            break;
                        case "END OF REPORT": //DeleteRows report end
                            LastRow -= DeleteRows(XlSh, CurrentRow - 3, CurrentRow);
                            CurrentRow -= 3;
                            break;
                        default:
                            XlSh.Cells[CurrentRow, 1].Value = ProjNo;
                            XlSh.Cells[CurrentRow, 2].Value = ProjName;
                            break;
                    }
                    CurrentRow += 1;
                }
                XlSh.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, XlSh.Cells[HeadingsRow, 1].CurrentRegion,false, 
                    Excel.XlYesNoGuess.xlYes).name ="Tab" + WBSTableName;
                xlAp.StatusBar = false;
                xlAp.ScreenUpdating = true;
                return ProjNo;
            }
            catch (Exception ex)
            {
                Globals.ThisAddIn.ExMsg(ex);
                return "";
            }
        }
        internal static string ParseArStatus(Excel.Worksheet XlSh) //return the project number or blank if an error occured
        {
            try
            {
                Excel.Application xlAp = Globals.ThisAddIn.Application;
                Excel.Workbook XlWb = xlAp.ActiveWorkbook;
                long CurrentRow; long LastRow;
                long HeadingsRow;
                string ProjNo = ""; string ProjName = "";
                string ArTableName;

                //Inititalise BST variables
                ProjNo = XlSh.Range["A44"].Text;
                ProjName = XlSh.Range["B39"].Text;

                //allocate table and sheet names
                xlAp.ScreenUpdating = false;
                ArTableName = "ArInv" + ProjNo;
                if (ExistSheet(XlWb, ArTableName))
                {
                    xlAp.DisplayAlerts = false;
                    XlWb.Worksheets[ArTableName].delete();
                    xlAp.DisplayAlerts = true;
                }
                XlSh.Name = ArTableName;

                //Parse report
                HeadingsRow = 41;
                LastRow = XlSh.UsedRange.Rows.Count;
                LastRow -= DeleteRows(XlSh, HeadingsRow + 1, HeadingsRow + 2); //Delete rows between headings and first table rows
                //LastRow -= DeleteRows(XlSh, 4, HeadingsRow - 1); //Delete rows between report name and headings
                //HeadingsRow = 5;
                //XlSh.Range[(HeadingsRow - 1) + ":" + (HeadingsRow - 1)].Insert(); //Insert blank row above headers

                XlSh.Range["B5"].TextToColumns(    //Split date
                    Type.Missing, //Destination
                    XlTextParsingType.xlDelimited, //DataType
                    XlTextQualifier.xlTextQualifierDoubleQuote,    //TextQualifier
                    false,         // Consecutive Delimiter
                    Type.Missing,  // Tab
                    Type.Missing,  // Semicolon
                    false,         // Comma
                    false,         // Space
                    true,          // Other
                    ":",           // Other Char
                    Type.Missing,  // Field Info
                    Type.Missing,  // Decimal Separator
                    Type.Missing,  // Thousands Separator
                    Type.Missing); // Trailing Minus Numbers
                SetNameRange(XlWb, "ImportDate"+ArTableName, XlSh.Range["C5"]);
                XlSh.Range["C5"].Value=XlSh.Range["C5"].Text.Trim();

                XlSh.Range[(HeadingsRow - 1) + ":" + (HeadingsRow - 1)].Clear();
                CurrentRow = HeadingsRow + 1; //First table row
                while (CurrentRow <= LastRow + 1)
                {
                    if (CurrentRow % 10 == 0)
                        xlAp.StatusBar = string.Format("Progress: {0:f0}%", CurrentRow * 100 / LastRow);

                    // Identify row type
                    string SwitchText = XlSh.Cells[CurrentRow, 1].Text.Trim();
                    switch (SwitchText)
                    {
                        case "A/R Status Report": //Delete page breaks
                            LastRow -= DeleteRows(XlSh, CurrentRow - 3, CurrentRow + 6);
                            CurrentRow -= 4; //3+1 because it is incremented by 1 later
                            break;
                        case "END OF REPORT": //DeleteRows report end
                            LastRow -= DeleteRows(XlSh, CurrentRow - 3, CurrentRow);
                            CurrentRow -= 3;
                            break;
                        default:
                            break;
                    }
                    CurrentRow += 1;
                }
                XlSh.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, XlSh.Cells[HeadingsRow, 1].CurrentRegion, false,
                    Excel.XlYesNoGuess.xlYes).name = "Tab" + ArTableName;
                xlAp.StatusBar = false;
                xlAp.ScreenUpdating = true;
                return ProjNo;
            }
            catch (Exception ex)
            {
                Globals.ThisAddIn.ExMsg(ex);
                return "";
            }
        }
        internal static void ParsePjAnalysis(Excel.Worksheet XlSh)
        {
            try
            {
                Excel.Application xlAp = Globals.ThisAddIn.Application;
                Excel.Workbook XlWb = xlAp.ActiveWorkbook;
                long CurrentRow = 0; long LastRow = 0;
                string AnalTableName = "PjAnalysis";
                long HeadingsRow;

                //allocate table and sheet names
                xlAp.ScreenUpdating = false;
                if (ExistSheet(XlWb, AnalTableName))
                {
                    xlAp.DisplayAlerts = false;
                    XlWb.Sheets[AnalTableName].Delete();
                    xlAp.DisplayAlerts = true;
                }
                XlSh.Name = AnalTableName;

                //Parse report
                HeadingsRow = 23;
                LastRow = XlSh.UsedRange.Rows.Count;
                LastRow -= DeleteRows(XlSh, HeadingsRow + 1, HeadingsRow + 2); //Delete rows between headings and first table rows
                                                                               //LastRow -= DeleteRows(XlSh, 4, HeadingsRow - 1); //Delete rows between report name and headings
                                                                               //HeadingsRow = 5;
                                                                               //XlSh.Range[(HeadingsRow - 1) + ":" + (HeadingsRow - 1)].Insert(); //Insert blank row above headers

                XlSh.Range["B5"].TextToColumns(    //Split date
                    Type.Missing, //Destination
                    XlTextParsingType.xlDelimited, //DataType
                    XlTextQualifier.xlTextQualifierDoubleQuote,    //TextQualifier
                    false,         // Consecutive Delimiter
                    Type.Missing,  // Tab
                    Type.Missing,  // Semicolon
                    false,         // Comma
                    false,         // Space
                    true,          // Other
                    ":",           // Other Char
                    Type.Missing,  // Field Info
                    Type.Missing,  // Decimal Separator
                    Type.Missing,  // Thousands Separator
                    Type.Missing); // Trailing Minus Numbers
                SetNameRange(XlWb, "ImportDate"+AnalTableName, XlSh.Range["C5"]);
                XlSh.Range["C5"].Value=XlSh.Range["C5"].Text.Trim();

                XlSh.Range[(HeadingsRow - 1) + ":" + (HeadingsRow - 1)].Clear();
                CurrentRow = HeadingsRow + 1; //First table row
                while (CurrentRow <= LastRow+1)
                {
                    if (CurrentRow % 10 == 0)
                        xlAp.StatusBar = string.Format("Progress: {0:f0}%", CurrentRow * 100 / LastRow);

                    // Identify row type
                    string SwitchText = XlSh.Cells[CurrentRow, 1].Text.Trim();
                    switch (SwitchText)
                    {
                        case "Project Analysis Report": //Delete page breaks
                            LastRow -= DeleteRows(XlSh, CurrentRow - 3, CurrentRow + 6);
                            CurrentRow -= 4; //3+1 because it is incremented by 1 later
                            break;
                        case "END OF REPORT": //DeleteRows report end
                            LastRow -= DeleteRows(XlSh, CurrentRow - 3, CurrentRow);
                            CurrentRow -= 3;
                            break;
                        default:
                            break;
                    }
                    CurrentRow += 1;
                }
                XlSh.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, XlSh.Cells[HeadingsRow, 1].CurrentRegion, false,
                    Excel.XlYesNoGuess.xlYes).name = "Tab" + AnalTableName;
                xlAp.StatusBar = false;
                xlAp.ScreenUpdating = true;
            }
            catch (Exception ex)
            {
                Globals.ThisAddIn.ExMsg(ex);
            }
        }
        internal static void ParseArAnalysis(Excel.Worksheet XlSh)
        {
            try
            {
                Excel.Application xlAp = Globals.ThisAddIn.Application;
                Excel.Workbook XlWb = xlAp.ActiveWorkbook;
                long CurrentRow = 0; long LastRow = 0;
                string AnalTableName;
                long HeadingsRow;

                //allocate table and sheet names
                xlAp.ScreenUpdating = false;
                AnalTableName = "ArAnalysis";
                if (ExistSheet(XlWb, AnalTableName))
                {
                    xlAp.DisplayAlerts = false;
                    XlWb.Sheets[AnalTableName].Delete();
                    xlAp.DisplayAlerts = true;
                }
                XlSh.Name = AnalTableName;

                //Parse report
                HeadingsRow = 35;
                LastRow = XlSh.UsedRange.Rows.Count;
                LastRow -= DeleteRows(XlSh, HeadingsRow + 1, HeadingsRow + 2); //Delete rows between headings and first table rows
                //LastRow -= DeleteRows(XlSh, 4, HeadingsRow - 1); //Delete rows between report name and headings
                //HeadingsRow = 5;
                //XlSh.Range[(HeadingsRow - 1) + ":" + (HeadingsRow - 1)].Insert(); //Insert blank row above headers

                XlSh.Range["B5"].TextToColumns(    //Split date
                    Type.Missing, //Destination
                    XlTextParsingType.xlDelimited, //DataType
                    XlTextQualifier.xlTextQualifierDoubleQuote,    //TextQualifier
                    false,         // Consecutive Delimiter
                    Type.Missing,  // Tab
                    Type.Missing,  // Semicolon
                    false,         // Comma
                    false,         // Space
                    true,          // Other
                    ":",           // Other Char
                    Type.Missing,  // Field Info
                    Type.Missing,  // Decimal Separator
                    Type.Missing,  // Thousands Separator
                    Type.Missing); // Trailing Minus Numbers
                SetNameRange(XlWb, "ImportDate"+AnalTableName, XlSh.Range["C5"]);
                XlSh.Range["C5"].Value=XlSh.Range["C5"].Text.Trim();

                XlSh.Range[(HeadingsRow - 1) + ":" + (HeadingsRow - 1)].Clear();
                CurrentRow = HeadingsRow + 1; //First table row
                while (CurrentRow <= LastRow + 1)
                {
                    if (CurrentRow % 10 == 0)
                        xlAp.StatusBar = string.Format("Progress: {0:f0}%", CurrentRow * 100 / LastRow);

                    // Identify row type
                    string SwitchText = XlSh.Cells[CurrentRow, 1].Text.Trim();
                    switch (SwitchText)
                    {
                        case "A/R Analysis Report": //Delete page breaks
                            LastRow -= DeleteRows(XlSh, CurrentRow - 3, CurrentRow + 6);
                            CurrentRow -= 4; //3+1 because it is incremented by 1 later
                            break;
                        case "END OF REPORT": //DeleteRows report end
                            LastRow -= DeleteRows(XlSh, CurrentRow - 3, CurrentRow);
                            CurrentRow -= 3;
                            break;
                        default:
                            break;
                    }
                    CurrentRow += 1;
                }
                XlSh.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, XlSh.Cells[HeadingsRow, 1].CurrentRegion, false,
                    Excel.XlYesNoGuess.xlYes).name = "Tab" + AnalTableName;
                xlAp.StatusBar = false;
                xlAp.ScreenUpdating = true;
            }
            catch (Exception ex)
            {
                Globals.ThisAddIn.ExMsg(ex);
            }
        }
        internal static void AddQueries(string[] TableNamesArray, string CombineQueryName, Excel.Workbook wbk) //Create queries for each table in TableNamesArray
        //https://stackoverflow.com/questions/61622872/adding-power-queries-to-excel-using-c-sharp
        //https://docs.microsoft.com/en-us/office/vba/language/reference/visual-basic-add-in-model/objects-visual-basic-add-in-model#vbcomponent
        //https://stackoverflow.com/questions/64210190/how-to-create-queries-and-connections#
        //http://www.cpearson.com/excel/vbe.aspx
        {
            try
            {
                string MacroName ;
                string wbkName = wbk.Name;
                string TableNames;
                VBComponent newStandardModule;
                string VBAcodeText;

/*                foreach (string TableName in TableNamesArray) //Return if query already exists
                {
                    if (ExistConnection(wbk, ConNamePrefix + TableName))
                    {
                        MessageBox.Show("Connection" + TableName + " already exists. Please delete existing connections");
                        return;
                    }
                }
*/
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
    CombineQueryName = ""{CombineQueryName}""
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
                List<string> WBSTables = new List<string>();

                //VBA project object model needs to be trusted for this to work
                if (VBATrusted(XlWb) == false) 
                {
                    MessageBox.Show("No Access to VB Project\n\rPlease allow access in Trusted Sources\n\r" +
                        "File > Options > Trust Center > Trust Center Settings > Macro Settings > Trust Access to the VBA project object model");
                    return;
                }

                //Collect all WBS table names and create queries
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
                if (WBSTables.Count<2) //Exit sub if there are one or less WBStables
                {
                    MessageBox.Show("Two or more WBS reports are required to create a WBS combined query");
                    return;
                } 
                xlAp.ScreenUpdating = false;
                AddQueries(WBSTables.ToArray(), "WBSCombineQuery", XlWb);

                xlAp.ScreenUpdating = true;
                MessageBox.Show("WBS combined query created");
            }
            catch (Exception ex)
            {
                Globals.ThisAddIn.ExMsg(ex);
            }
        }
        internal static void AddPMCol() //Add PM columns to PjAnalysis report
        {
            try
            {
                Excel.Application xlAp = Globals.ThisAddIn.Application;
                Excel.Workbook XlWb = xlAp.ActiveWorkbook;
                Excel.ListObject PjAnalysisTable;
                Excel.ListColumn LCol;
                string PjAnalTableName = "TabPjAnalysis";
                string PjAnaShtName;
                string PMCol1 = "Adjustments";
                string PMCol2 = "NewBudgert";
                string PMCol3 = "New % Complete (Get calculation from PM)";
                string PMCol4 = "New Revenue generated for the Project Lifetime";
                string PMCol5 = "Revenue for the MONTH (If the TOTAL is a Negative Revenue, a WIP write off form is needed.)";
                string PMCol6 = "Double check where it states true";

                if ((PjAnaShtName = ExistListObject(PjAnalTableName)) != "")
                {
                    PjAnalysisTable = XlWb.Worksheets[PjAnaShtName].ListObjects[PjAnalTableName];
                    if (!ExistListCol(PjAnalysisTable,PMCol1))
                        {
                        LCol = PjAnalysisTable.ListColumns.Add(); //Adjustments
                        LCol.Name = PMCol1;

                        LCol = PjAnalysisTable.ListColumns.Add(); //New budget
                        LCol.Name = PMCol2;
                        LCol.DataBodyRange.Value = "=[@[Tot Bdgt Eff]]+[@[" + PMCol1 + "]]";

                        LCol = PjAnalysisTable.ListColumns.Add(); //New % complete
                        LCol.Name = PMCol3;

                        LCol = PjAnalysisTable.ListColumns.Add(); //New revenue
                        LCol.Name = PMCol4;
                        LCol.DataBodyRange.Value = "=[@[" + PMCol2 + "]]*[@[" + PMCol3 + "]]/100";

                        LCol = PjAnalysisTable.ListColumns.Add(); //Revenue/ month
                        LCol.Name = PMCol5;
                        LCol.DataBodyRange.Value = "=[@[" + PMCol4 + "]]-[@Revenue]";

                        LCol = PjAnalysisTable.ListColumns.Add(); //Check
                        LCol.Name = PMCol6;
                        LCol.DataBodyRange.Value = "=[@[" + PMCol5 + "]]<-1";
                        LCol.DataBodyRange.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlEqual, "=TRUE");
                        LCol.DataBodyRange.FormatConditions[1].Interior.Color = 13551615;

                        PjAnalysisTable.HeaderRowRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    }
                    else
                    {
                        MessageBox.Show("PM columns are already added");
                        return;
                    }
                }
                else
                {
                    MessageBox.Show("Project analysis table does not exist");
                    return;
                }

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
        internal static string ExistListObject(string ListName) // Returns sheet name if the list object exist
        {
            Excel.Application xlAp = Globals.ThisAddIn.Application;
            Excel.Workbook XlWb = xlAp.ActiveWorkbook;

            foreach (Excel.Worksheet Sheet in XlWb.Worksheets) // Loop through all the worksheets
            {
                foreach (Excel.ListObject ListObj in Sheet.ListObjects) // Loop through each table in the worksheet
                {
                    if (ListObj.Name == ListName)
                    {
                        return Sheet.Name;
                    }
                }
            }

            return "";
        }
        internal static bool ExistListCol(Excel.ListObject ListObj, string ColName) // Returns true if column name exists in the listobjet
        {
            Excel.Application xlAp = Globals.ThisAddIn.Application;
            Excel.Workbook XlWb = xlAp.ActiveWorkbook;

            foreach (Excel.ListColumn LCol in ListObj.ListColumns) // Loop through all the columns
            {
                     if (LCol.Name == ColName)
                    {
                        return true;
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
        internal static void SetNameRange(Excel.Workbook Wb, string ChkName, Excel.Range NameRange)
        {           
            foreach(Excel.Name DefName in Wb.Names)
            {
                if (DefName.Name == ChkName)
                {
                   DefName.Delete();
                }
           }        
            NameRange.Name=ChkName;
         }
    }
}
