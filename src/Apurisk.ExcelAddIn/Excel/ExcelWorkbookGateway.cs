using System;
using System.Runtime.InteropServices;
namespace Apurisk.ExcelAddIn.Excel
{
    public sealed class ExcelWorkbookGateway
    {
        private readonly object _excelApplication;

        public ExcelWorkbookGateway(object excelApplication)
        {
            _excelApplication = excelApplication;
        }

        public void EnsureSheet(string sheetName, string[] headers)
        {
            dynamic excel = _excelApplication;
            dynamic workbook = excel.ActiveWorkbook;

            if (workbook == null)
            {
                workbook = excel.Workbooks.Add();
            }

            EnsureSheetInternal(workbook, sheetName, headers);
        }

        public void ActivateSheet(string sheetName)
        {
            dynamic excel = _excelApplication;
            dynamic sheet = FindSheet(excel.ActiveWorkbook, sheetName);
            if (sheet != null)
            {
                sheet.Activate();
            }
        }

        private static void EnsureSheetInternal(dynamic workbook, string sheetName, string[] headers)
        {
            dynamic sheet = FindSheet(workbook, sheetName);
            if (sheet == null)
            {
                sheet = workbook.Worksheets.Add();
                sheet.Name = sheetName;
            }

            for (int i = 0; i < headers.Length; i++)
            {
                dynamic cell = sheet.Cells[1, i + 1];
                cell.Value2 = headers[i];
                cell.Font.Bold = true;
                Marshal.ReleaseComObject(cell);
            }

            dynamic usedRange = sheet.UsedRange;
            usedRange.Columns.AutoFit();
            Marshal.ReleaseComObject(usedRange);
        }

        private static dynamic FindSheet(dynamic workbook, string sheetName)
        {
            if (workbook == null)
            {
                return null;
            }

            foreach (dynamic sheet in workbook.Worksheets)
            {
                string currentName = sheet.Name as string;
                if (string.Equals(currentName, sheetName, StringComparison.OrdinalIgnoreCase))
                {
                    return sheet;
                }

                Marshal.ReleaseComObject(sheet);
            }

            return null;
        }
    }
}
