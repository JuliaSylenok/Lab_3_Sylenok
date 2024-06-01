using System;
using System.Collections.Generic;
using System.Reflection;

namespace lab_3
{
    class Program
    {
        static void Main(string[] args)
        {
            List<string> appointments = new List<string>
            {
                "Іван Петров, Стоматологія, Пломбування, 2024-05-15, 11:00",
                "Анна Іванова, Косметологія, Чистка обличчя, 2024-05-20, 14:30",
                "Олександр Сидоров, Дерматологія, Консультація, 2024-06-01, 10:00"
            };

            Assembly wordAssembly = Assembly.LoadFrom(@"C:\VARIANT 19\2 course\АППЗ\Lab_3_Sylenok\WordExport\bin\Debug\net6.0\WordExport.dll");
            Assembly excelAssembly = Assembly.LoadFrom(@"C:\VARIANT 19\2 course\АППЗ\Lab_3_Sylenok\ExcelExport\bin\Debug\net6.0\ExcelExport.dll");

            Type wordType = wordAssembly.GetType("WordExport.WordExporter");
            Type excelType = excelAssembly.GetType("ExcelExport.ExcelExporter");

            dynamic wordExporter = Activator.CreateInstance(wordType);
            dynamic excelExporter = Activator.CreateInstance(excelType);

            string wordPath = @"C:\VARIANT 19\2 course\АППЗ\Lab_3_Sylenok\AppointmentReport.doc";
            string excelPath = @"C:\VARIANT 19\2 course\АППЗ\Lab_3_Sylenok\AppointmentReport.xlsx";

            wordExporter.ExportToWord(appointments, wordPath);
            excelExporter.ExportToExcel(appointments, excelPath);

            Console.WriteLine("Appointment reports have been generated.");
        }
    }
}
