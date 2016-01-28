using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RF
{
    public class ExcelHelper
    {
        public static void SaveReport(List<Item> items, string filepath)
        {
            if (items == null || items.Count == 0) return;
            var monthString = DateTime.Today.Month < 10 ? "0" + DateTime.Today.Month : DateTime.Today.Month + string.Empty;
            var dayString = DateTime.Today.Day < 10 ? "0" + DateTime.Today.Day : DateTime.Today.Day + string.Empty;
            var reportString = string.Format("Report-{0}{1}{2}-{3}{4}", DateTime.Today.Year, monthString, dayString, DateTime.Now.Hour, DateTime.Now.Minute);
            var misValue = System.Reflection.Missing.Value;


            var xlApp = new Microsoft.Office.Interop.Excel.Application();

            var xlWorkBook = xlApp.Workbooks.Add(misValue);
            var xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.Item[1];
            int i;

            //Add data.
            xlWorkSheet.Cells[1, 1] = "#";
            xlWorkSheet.Cells[1, 2] = "Line number";
            xlWorkSheet.Cells[1, 3] = "Original string";
            xlWorkSheet.Cells[1, 4] = "Extracted value";
            xlWorkSheet.Cells[1, 5] = "Suggestion";
            xlWorkSheet.Cells[1, 6] = "File Path";
            xlWorkSheet.Cells[1, 7] = "Need change";
            xlWorkSheet.Cells[1, 8] = "Reviewd";

            var titlerange = xlWorkSheet.Range["a1", "g1"];
            titlerange.Font.Bold = true;
            titlerange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            titlerange.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            int index = 1;

            foreach (var item in items)
            {
                var rowIndex = index + 1;
                xlWorkSheet.Cells[rowIndex, 1] = index;
                xlWorkSheet.Cells[rowIndex, 2] = item.Line;
                xlWorkSheet.Cells[rowIndex, 3] = item.Original;
                xlWorkSheet.Cells[rowIndex, 4] = item.ExtractedValue;
                xlWorkSheet.Cells[rowIndex, 5] = item.Suggestion;
                xlWorkSheet.Cells[rowIndex, 6] = item.FilePath;
                xlWorkSheet.Cells[rowIndex, 7] = "---";
                xlWorkSheet.Cells[rowIndex, 8] = "---";
                index++;
            }


            xlWorkSheet.Columns.AutoFit();
            xlWorkSheet.Rows.AutoFit();

            xlApp.Visible = true;
            xlWorkBook.SaveAs(filepath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
        }

    }
}
