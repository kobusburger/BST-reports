using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace BST_reports
{
    static class Pivot
    {
        internal static void AddHoursTable(string Tablename)
        {
            // Add pivot table
            Excel.Worksheet XlSh;
            Excel.Workbook XlWb;
            Excel.Application XlAp;
            Excel.PivotCache PCache;
            Excel.PivotTable PTable;
            Excel.PivotField PField;
            try
            {
                XlAp = Globals.ThisAddIn.Application;
                XlWb = XlAp.ActiveWorkbook;
                PCache = XlWb.PivotCaches().Create(Excel.XlPivotTableSourceType.xlDatabase, Tablename);

                // Add pivot table
                XlSh = XlWb.Sheets.Add();
                PTable = XlSh.PivotTables().Add(PCache, XlSh.Range["A1"]);
                PTable.AddFields("name", "Month", "Cost Type");
                PTable.AddDataField(Field: PTable.PivotFields("Hours / Quantity"), Function: Excel.XlConsolidationFunction.xlSum);
                PTable.ClearAllFilters();
                PField = PTable.PivotFields("Cost Type");
                PField.CurrentPage = "Labor";
            }
            catch (Exception ex)
            {
                ExMsg(ex);
            }
        }

        internal static void AddCostChart(string Tablename)
        {
            // Add pivot table and chart
            Excel.Worksheet XlSh;
            Excel.Workbook XlWb;
            Excel.Application xlAp;
            Excel.PivotCache PCache;
            Excel.PivotTable PTable;
            Excel.PivotField PField;
            Excel.Shape PChart;
            try
            {
                xlAp = Globals.ThisAddIn.Application;
                XlWb = xlAp.ActiveWorkbook;
                PCache = XlWb.PivotCaches().Create(Excel.XlPivotTableSourceType.xlDatabase, Tablename);

                // Add pivot table
                XlSh = XlWb.Sheets.Add();
                PTable = XlSh.PivotTables().Add(PCache, XlSh.Range["A1"]);
                PTable.AddFields("Month", new[] { "Phase", "Task" }, new[] { "Project", "Cost Type" }); // (RowFields, ColumnFields, PageFields)
                {
                    var withBlock = PTable.AddDataField(PTable.PivotFields("Cost Amount"));
                    withBlock.Function = Excel.XlConsolidationFunction.xlSum;
                    withBlock.Calculation = Excel.XlPivotFieldCalculation.xlRunningTotal;
                    withBlock.NumberFormat = "#,##0"; // The comma does not mean that the thousand separator is a comma. I means that the locale thousand separator should be used
                    withBlock.BaseField = "Month";
                }

                PChart = XlSh.Shapes.AddChart2(-1, Excel.XlChartType.xlLine, 10, 10, 800, 400);
                PTable.ClearAllFilters();
                PField = PTable.PivotFields("Cost Type");
            }
            catch (Exception ex)
            {
                ExMsg(ex);
            }
        }

        internal static bool ExistListObject(string ListName)
        {
            bool ExistListObjectRet = default;
            // Returns true if a list object exist in the active workbook
            Excel.Application xlAp;
            Excel.Workbook XlWb;
            xlAp = Globals.ThisAddIn.Application;
            XlWb = xlAp.ActiveWorkbook;
            ExistListObjectRet = false;
            foreach (Excel.Worksheet xlWs in XlWb.Worksheets) // Loop through all the worksheets
            {
                foreach (Excel.ListObject ListObj in xlWs.ListObjects) // Loop through each table in the worksheet
                {
                    if (ListObj.Name == ListName)
                    {
                        return true;
                    }
                }
            }

            return ExistListObjectRet;
        }

        static void ExMsg(Exception Ex)
        {
            MessageBox.Show(Ex.ToString(), "BST Add-In exception (copy text with Ctrl+C)");
        }

    }
}
