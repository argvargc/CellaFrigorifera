using System;
using System.Drawing;
using ClosedXML.Excel; //Libreria che deve essere obbligatoriamente installata con "NuGet"
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Spreadsheet;

class Program
{
    static void Main()
    {
        string filePath = "C:/Documenti/CellaFrigorifera.xlsx";

        // 1. CREAZIONE E SCRITTURA DEL FILE EXCEL
        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.Worksheets.Add("CellaFrigorifera");

            // Intestazioni
            int nibbleInput = 0;

            worksheet.Cell(1, 1).Value = "Rilevazione";
            worksheet.Cell(1, 2).Value = "Nibble";
            worksheet.Cell(1, 3).Value = "Hex";
            worksheet.Cell(1, 4).Value = "Note";

            for (int i = 1; i < 6; i++)
            {
                Console.WriteLine("Inserisci valore binario per nibble " + i);
                nibbleInput = Convert.ToInt16(Console.ReadLine());
                worksheet.Cell(i + 1, 2).Value = nibbleInput;
                worksheet.Cell(i + 1, 1).Value = "Rilevazione " + i;
                if (nibbleInput > 110)
                {
                    worksheet.Cell(i + 1, 4).Value = "Attenzione! Possibile avaria! ";
                    var noteRow = worksheet.Row(i + 1);
                    noteRow.Style.Fill.BackgroundColor = XLColor.Red;
                    noteRow.Style.Font.Bold = true;
                }
                else if (nibbleInput <= 2)
                {
                    worksheet.Cell(i + 1, 4).Value = "Attenzione! Rischio congelamento!  ";
                }
                else
                {
                    worksheet.Cell(i + 1, 4).Value = "Temperatura nella norma";
                }
            }

            // Formattazione intestazione
            var headerRow = worksheet.Row(1);
            headerRow.Style.Font.Bold = true; //Grassetto
            headerRow.Style.Fill.BackgroundColor = XLColor.Yellow; //Colore Sfondo Cella
            headerRow.Style.Font.FontColor = XLColor.Blue; //Colore del Testo

            worksheet.Cell(2, 3).FormulaA1 = "=BINARIO.HEX(B2)"; //Cella con Formula Calcolata
            worksheet.Cell(3, 3).FormulaA1 = "=BINARIO.HEX(B3)"; //Cella con Formula Calcolata
            worksheet.Cell(4, 3).FormulaA1 = "=BINARIO.HEX(B4)"; //Cella con Formula Calcolata
            worksheet.Cell(5, 3).FormulaA1 = "=BINARIO.HEX(B5)"; //Cella con Formula Calcolata
            worksheet.Cell(6, 3).FormulaA1 = "=BINARIO.HEX(B6)"; //Cella con Formula Calcolata


            // Adatta la larghezza delle colonne al contenuto
            worksheet.Columns().AdjustToContents();
            worksheet.Column("A").Width = 16;
            worksheet.Columns("B:J").Width = worksheet.Column("A").Width;

            // Salva il file
            workbook.SaveAs(filePath);
            Console.WriteLine($"\nIl File '{filePath}' e' stato creato con successo!\n");
        }

        // 2. LETTURA DEL FILE EXCEL
        using (var workbook = new XLWorkbook(filePath))
        {
            var worksheet = workbook.Worksheet("CellaFrigorifera");

            Console.WriteLine("Contenuto del file:");
            foreach (var row in worksheet.RowsUsed())
            {
                string s1 = row.Cell(1).GetString();
                string s2 = row.Cell(2).GetString();
                string s3 = row.Cell(3).GetString();
                string s4 = row.Cell(4).GetString();
                string s5 = row.Cell(5).GetString();
                string s6 = row.Cell(6).GetString();
                string s7 = row.Cell(7).GetString();
                string s8 = row.Cell(8).GetString();
                string s9 = row.Cell(9).GetString();
                string s10 = row.Cell(10).GetString();
                //Console.WriteLine($"{id}\t| {nome}\t| {eta}\t| {biglie}");
                Console.WriteLine(s1 + " | " + s2 + " | " + s3 + " | " + s4 + " | " + s5);
            }
        }
    }
}
