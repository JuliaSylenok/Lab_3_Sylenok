using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace WordExport
{
    public class WordExporter
    {
        public void ExportToWord(List<string> appointments, string wordPath)
        {
            Word.Application wdApp = new Word.Application();
            Word.Document doc = wdApp.Documents.Add();
            CustomizeWordDocument(doc, appointments);
            doc.SaveAs2(wordPath);
            doc.Close();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
            wdApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wdApp);
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void CustomizeWordDocument(Word.Document doc, List<string> appointments)
        {
            Word.Paragraph header = doc.Paragraphs.Add();
            header.Range.Text = "Appointment Report";
            header.Range.Font.Size = 24;
            header.Range.Font.Name = "Times New Roman";
            header.Range.InsertParagraphAfter();

            Word.Paragraph appointmentsHeader = doc.Paragraphs.Add();
            appointmentsHeader.Range.Text = "Scheduled Appointments";
            appointmentsHeader.Range.Font.Name = "Times New Roman";
            appointmentsHeader.Range.Font.Size = 14;
            appointmentsHeader.Range.InsertParagraphAfter();

            Word.Table appointmentsTable = doc.Tables.Add(appointmentsHeader.Range, appointments.Count + 1, 5);
            appointmentsTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            appointmentsTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

            appointmentsTable.Cell(1, 1).Range.Text = "Ім'я";
            appointmentsTable.Cell(1, 2).Range.Text = "Категорія";
            appointmentsTable.Cell(1, 3).Range.Text = "Опис";
            appointmentsTable.Cell(1, 4).Range.Text = "Дата";
            appointmentsTable.Cell(1, 5).Range.Text = "Час";

            for (int i = 0; i < appointments.Count; i++)
            {
                string[] parts = appointments[i].Split(',');

                appointmentsTable.Cell(i + 2, 1).Range.Text = parts[0].Trim();
                appointmentsTable.Cell(i + 2, 2).Range.Text = parts[1].Trim();
                appointmentsTable.Cell(i + 2, 3).Range.Text = parts[2].Trim();
                appointmentsTable.Cell(i + 2, 4).Range.Text = parts[3].Trim();
                appointmentsTable.Cell(i + 2, 5).Range.Text = parts[4].Trim();
            }
        }
    }
}
