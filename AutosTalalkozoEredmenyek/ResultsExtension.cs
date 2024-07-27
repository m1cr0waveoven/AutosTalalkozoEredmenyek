using AutosTalalkozoEredmenyek.Models;
using ClosedXML.Excel;
using System;
using System.Diagnostics;
using System.IO;

namespace AutosTalalkozoEredmenyek;

internal static class ResultsExtension
{
    public static void WriteToExcel(this Results results)
    {
        if (results is null)
            return;

        try
        {
            string workbookPath = Path.Combine(Directory.GetCurrentDirectory(), "Eredmenyek.xlsx");
            using var workbook = new XLWorkbook(workbookPath);

            const int rowsToSkip = 3;
            var worksheet = workbook.Worksheets.Worksheet("Autolimbó");

            worksheet.Cells("A3:E100").Clear();
            var autolimbo_results = results.Autolimbo.Results;
            int gyariIndex = 0;
            int epitettIndex = 0;
            for (int i = 0; i < autolimbo_results.Count; i++)
            {
                if (autolimbo_results[i] is { Kategoria: "gyári" })
                {
                    worksheet.Cell(gyariIndex + rowsToSkip, 1).Value = autolimbo_results[i].Rendszam;
                    worksheet.Cell(gyariIndex + rowsToSkip, 2).Value = autolimbo_results[i].Magassag;
                    gyariIndex++;
                    continue;
                }

                worksheet.Cell(epitettIndex + rowsToSkip, 4).Value = autolimbo_results[i].Rendszam;
                worksheet.Cell(epitettIndex + rowsToSkip, 5).Value = autolimbo_results[i].Magassag;
                epitettIndex++;
            }
            worksheet.Columns().AdjustToContents();

            //Időmérő eredmények 
            worksheet = workbook.Worksheets.Worksheet("Időmérő");
            worksheet.Cells("A3:Y100").Clear();
            int row;
            var felnikitartas_no = results.FelnikitartasNo.Results;
            for (int i = 0; i < felnikitartas_no.Count; i++)
            {
                row = i + rowsToSkip;
                worksheet.Cell(row, 1).Value = felnikitartas_no[i].Nev;
                worksheet.Cell(row, 2).Value = felnikitartas_no[i].Ido;
            }

            var felnikitartas_ferfi = results.FelnikitartasFerfi.Results;
            for (int i = 0; i < felnikitartas_ferfi.Count; i++)
            {
                row = i + rowsToSkip;
                worksheet.Cell(row, 4).Value = felnikitartas_ferfi[i].Nev;
                worksheet.Cell(row, 5).Value = felnikitartas_ferfi[i].Ido;
            }
            //worksheet.Columns().AdjustToContents();

            var gumiguritas = results.Gumiguritas.Results;
            for (int i = 0; i < gumiguritas.Count; i++)
            {
                row = i + rowsToSkip;
                worksheet.Cell(row, 7).Value = gumiguritas[i].Nev;
                worksheet.Cell(row, 8).Value = gumiguritas[i].Ido;
                worksheet.Cell(row, 9).Value = gumiguritas[i].Hibapont;
                worksheet.Cell(row, 10).FormulaR1C1 = "=H" + row + "+I" + row + "*K1";
            }

            var autoszlalom = results.Autoszlalom.Results;
            for (int i = 0; i < autoszlalom.Count; i++)
            {
                row = i + rowsToSkip;
                worksheet.Cell(row, 12).Value = autoszlalom[i].Nev;
                worksheet.Cell(row, 13).Value = autoszlalom[i].Ido;
                worksheet.Cell(row, 14).Value = autoszlalom[i].Hibapont;
                worksheet.Cell(row, 15).FormulaR1C1 = "=M" + row + "+N" + row + "*P1";
            }

            var kviz = results.Kviz.Results;
            for (int i = 0; i < kviz.Count; i++)
            {
                row = i + rowsToSkip;
                worksheet.Cell(row, 17).Value = kviz[i].Nev;
                worksheet.Cell(row, 18).Value = kviz[i].Ido;
                worksheet.Cell(row, 19).Value = kviz[i].Hibapont;
                worksheet.Cell(row, 20).FormulaR1C1 = "=R" + row + "+S" + row + "*U1";
            }

            var autoToloHuzo = results.AutoToloHuzo.Results;
            for (int i = 0; i < autoToloHuzo.Count; i++)
            {
                row = i + rowsToSkip;
                worksheet.Cell(row, 22).Value = autoToloHuzo[i].Nev;
                worksheet.Cell(row, 23).Value = autoToloHuzo[i].Ido;
                worksheet.Cell(row, 24).Value = autoToloHuzo[i].Hibapont;
                worksheet.Cell(row, 25).FormulaR1C1 = "=W" + row + "+X" + row + "*Z1";
            }

            var kipufogohangyomas = results.KipufogoHangnyomas.Results;
            worksheet = workbook.Worksheets.Worksheet("Kipufogó hangnyomás");
            worksheet.Cells("A3:I100").Clear();
            int dizelIndex = 0;
            int benzinIndex = 0;

            for (int i = 0; i < kipufogohangyomas.Count; i++)
            {
                if (kipufogohangyomas[i] is { MotorTipus: "dizel" })
                {
                    worksheet.Cell(dizelIndex + rowsToSkip, 1).Value = kipufogohangyomas[i].Rendszam;
                    worksheet.Cell(dizelIndex + rowsToSkip, 2).Value = kipufogohangyomas[i].MotorTipus;
                    worksheet.Cell(dizelIndex + rowsToSkip, 3).Value = kipufogohangyomas[i].Hangnyomas;
                    dizelIndex++;
                    continue;
                }

                worksheet.Cell(benzinIndex + rowsToSkip, 5).Value = kipufogohangyomas[i].Rendszam;
                worksheet.Cell(benzinIndex + rowsToSkip, 6).Value = kipufogohangyomas[i].MotorTipus;
                worksheet.Cell(benzinIndex + rowsToSkip, 7).Value = kipufogohangyomas[i].Hangnyomas;
                benzinIndex++;

            }

            //Autoszépségverseny
            worksheet = workbook.Worksheets.Worksheet("Autoszépségverseny");
            worksheet.Cells("A3:G100").Clear();
            var autoszepsegverseny = results.Autoszepsegverseny.Results;
            for (int i = 0; i < autoszepsegverseny.Count; i++)
            {
                row = i + rowsToSkip;
                worksheet.Cell(row, 1).Value = autoszepsegverseny[i].Rendszam;
                worksheet.Cell(row, 2).Value = autoszepsegverseny[i].Kulso;
                worksheet.Cell(row, 3).Value = autoszepsegverseny[i].Belso;
                worksheet.Cell(row, 4).Value = autoszepsegverseny[i].Motorter;
                worksheet.Cell(row, 5).Value = autoszepsegverseny[i].Felni;
                worksheet.Cell(row, 6).Value = autoszepsegverseny[i].Osszhang;
                worksheet.Cell(row, 7).Value = autoszepsegverseny[i].Osszpontszam;
            }

            workbook.Save();
            Process.Start(workbookPath);

        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
            Console.ReadKey();
        }
    }
}
