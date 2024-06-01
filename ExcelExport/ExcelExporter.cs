using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelExport
{
    public class ExcelExporter
    {
        public void ExportToExcel(List<string> appointments, string excelPath)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlBook = xlApp.Workbooks.Add();
            Excel.Worksheet xlSheet1 = xlBook.Sheets[1];
            xlSheet1.Name = "Appointments";
            CustomizeExcelWorkbook(xlSheet1, appointments);
            xlBook.SaveAs(excelPath);
            xlBook.Close();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheet1);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBook);
            xlApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void CustomizeExcelWorkbook(Excel.Worksheet xlSheet1, List<string> appointments)
        {
            xlSheet1.Cells[1, 1].Value = "Ім'я";
            xlSheet1.Cells[1, 2].Value = "Категорія";
            xlSheet1.Cells[1, 3].Value = "Опис";
            xlSheet1.Cells[1, 4].Value = "Дата";
            xlSheet1.Cells[1, 5].Value = "Час";

            for (int i = 0; i < appointments.Count; i++)
            {
                string[] parts = appointments[i].Split(',');

                xlSheet1.Cells[i + 2, 1].Value = parts[0].Trim();
                xlSheet1.Cells[i + 2, 2].Value = parts[1].Trim();
                xlSheet1.Cells[i + 2, 3].Value = parts[2].Trim();
                xlSheet1.Cells[i + 2, 4].Value = parts[3].Trim();
                xlSheet1.Cells[i + 2, 5].Value = parts[4].Trim();
            }

            Excel.Range tableRange = xlSheet1.Range["A1", xlSheet1.Cells[appointments.Count + 1, 5]];
            tableRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            tableRange.Font.Name = "Times New Roman";
            tableRange.Font.Size = 14;
            tableRange.Columns.AutoFit();
        }
    }
}
