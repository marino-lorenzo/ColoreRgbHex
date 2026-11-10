using System;
using System.Drawing;
using ClosedXML.Excel; //Libreria che deve essere obbligatoriamente installata con "NuGet"
using DocumentFormat.OpenXml.Spreadsheet;

class Program
{
    static void Main()
    {
        string filePath = "C:/Users/marino.lorenzo/Documents/vs projects/c#/ColoreRgbHex/Risultato.xlsx";

        // 1. CREAZIONE E SCRITTURA DEL FILE EXCEL
        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.Worksheets.Add("Utenti");

            int R, G, B;

            do
            {
                Console.WriteLine("Inserire valore decimale del Rosso(0-255)");
                R = Convert.ToInt16(Console.ReadLine());
                if (R < 0 || R > 255)
                {
                    Console.WriteLine("Si possono inserire solo valori compresi tra 0 e 255, riprovare");
                }
            } while (R < 0 || R > 255);
            do
            {
                Console.WriteLine("Inserire valore decimale del Verde(0-255)");
                G = Convert.ToInt16(Console.ReadLine());
                if (G < 0 || G > 255)
                {
                    Console.WriteLine("Si possono inserire solo valori compresi tra 0 e 255, riprovare");
                }
            } while (G < 0 || G > 255);
            do
            {
                Console.WriteLine("Inserire valore decimale del Blu(0-255)");
                B = Convert.ToInt16(Console.ReadLine());
                if (B < 0 || B > 255)
                {
                    Console.WriteLine("Si possono inserire solo valori compresi tra 0 e 255, riprovare");
                }
            } while (B < 0 || B > 255);

            worksheet.Cell("A1").Value = "RGB(" + R + "," + G + "," + B + ")";
            worksheet.Cell("A1").Style.Fill.BackgroundColor = XLColor.FromArgb(R, G, B);
            worksheet.Column(1).AdjustToContents();

            workbook.SaveAs(filePath);
            Console.WriteLine($"\nIl File '{filePath}' e' stato creato con successo!\n");
        }
    }
}
