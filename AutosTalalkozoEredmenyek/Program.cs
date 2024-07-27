using AutosTalalkozoEredmenyek.Models;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace AutosTalalkozoEredmenyek;
// Created in 2020
// Refactored in 2024
internal class Program
{
    private static readonly IConfiguration Configuration = new ConfigurationBuilder()
           .SetBasePath(Directory.GetCurrentDirectory())
           .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
           .AddUserSecrets<Program>()
           .Build();
    static void Main(string[] args)
    {
        var results = new Results();
        results.Autolimbo = GetResults<Autolimbo>(Configuration.GetValue<string>("AutolimboUri"));
        results.FelnikitartasFerfi = GetResults<Felnikitartas>(Configuration.GetValue<string>("FelnikitartasFerfiUri"));
        results.FelnikitartasNo = GetResults<Felnikitartas>(Configuration.GetValue<string>("FelnikitartasNoUri"));
        results.Autoszepsegverseny = GetResults<Autoszepsegverseny>(Configuration.GetValue<string>("AutoszepsegversenyUri"));
        results.Gumiguritas = GetResults<Idomero>(Configuration.GetValue<string>("GumikitartasUri"));
        results.Autoszlalom = GetResults<Idomero>(Configuration.GetValue<string>("AutoszlalomUri"));
        results.Kviz = GetResults<Idomero>(Configuration.GetValue<string>("KvizUri"));
        results.AutoToloHuzo = GetResults<Idomero>(Configuration.GetValue<string>("AutoToloHuzoUri"));
        results.KipufogoHangnyomas = GetResults<Kipufogohangyomas>(Configuration.GetValue<string>("KipufogohangnyomasUri"));

        results.WriteToExcel();
    }

    private static IResultModel<T> GetResults<T>(string path)
    {
        if (string.IsNullOrEmpty(path))
            return new NoResultModel<T> { Error = "Üres elérési út.", Message = "Érvénytelen elérési út, az adatokat nem lehet lekérdezni." };

        try
        {
            HttpWebRequest WebReq = CreateWebRequest(path);

            HttpWebResponse response = (HttpWebResponse)WebReq.GetResponse();
            string jsonString = string.Empty;
            using (Stream stream = response.GetResponseStream())
            using (var reader = new StreamReader(stream, System.Text.Encoding.UTF8))
            {
                jsonString = reader.ReadToEnd();
            }

            // JObject jsonResponse = JObject.Parse(jsonString);
            // JObject error = (JObject)error["error"];
            // JArray results = (JArray)jsonResponse["results"];

            var resut = JsonConvert.DeserializeObject<ResultModel<T>>(jsonString);

            return resut;
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
            Console.ReadKey();
            return new NoResultModel<T> { Error = ex.Message };
        }

    }
    private static HttpWebRequest CreateWebRequest(string path)
    {
        HttpWebRequest webReq = (HttpWebRequest)WebRequest.Create(path);
        webReq.Credentials = new NetworkCredential(Configuration.GetValue<string>("userName"), Configuration.GetValue<string>("password"));
        webReq.Method = "GET";
        return webReq;
    }

    [Obsolete("Use WriteToExcel extension method on Result instead")]
    private static void WriteToExcel(List<Autolimbo> autolimbo_results)
    {
        Excel.Application oXL = null;
        Excel._Workbook oWB = null;
        Excel._Worksheet oSheet;
        Excel.Range oRng;

        var missingValue = System.Reflection.Missing.Value;
        try
        {
            // Start Excel and get Application object.
            oXL = new Excel.Application();
            oXL.Visible = true;

            // Get a new workbook.
            string path = Path.Combine(Directory.GetCurrentDirectory(), "Eredmenyek.xlsx");
            oWB = (Excel._Workbook)oXL.Workbooks.Open(path);
            oSheet = (Excel._Worksheet)oWB.Sheets[1];
            for (int i = 0; i < autolimbo_results.Count; i++)
            {
                oSheet.Cells[i + 2, 1] = autolimbo_results[i].Rendszam;
                oSheet.Cells[i + 2, 2] = autolimbo_results[i].Kategoria;
                oSheet.Cells[i + 2, 3] = autolimbo_results[i].Magassag;
            }
            #region unused_code
            ////Add table headers going cell by cell.
            //oSheet.Cells[1, 1] = "First Name";
            //oSheet.Cells[1, 2] = "Last Name";
            //oSheet.Cells[1, 3] = "Full Name";
            //oSheet.Cells[1, 4] = "Salary";

            //Format A1:D1 as bold, vertical alignment = center.
            //oSheet.get_Range("A1", "D1").Font.Bold = true;
            //oSheet.get_Range("A1", "D1").VerticalAlignment =
            //    Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

            // Create an array to multiple values at once.
            //string[,] saNames = new string[5, 2];

            //saNames[0, 0] = "John";
            //saNames[0, 1] = "Smith";
            //saNames[1, 0] = "Tom";

            //saNames[4, 1] = "Johnson";

            //Fill A2:B6 with an array of values (First and Last Names).
            //oSheet.get_Range("A2", "B6").Value2 = saNames;

            ////Fill C2:C6 with a relative formula (=A2 & " " & B2).
            //oRng = oSheet.get_Range("C2", "C6");
            //oRng.Formula = "=A2 & \" \" & B2";

            ////Fill D2:D6 with a formula(=RAND()*100000) and apply format.
            //oRng = oSheet.get_Range("D2", "D6");
            //oRng.Formula = "=RAND()*100000";
            //oRng.NumberFormat = "$0.00";
            #endregion
            //AutoFit columns A:D.
            oRng = oSheet.get_Range("A1", "D1");
            oRng.EntireColumn.AutoFit();

            oXL.Visible = false;
            oXL.UserControl = false;
            oWB.SaveAs(path, Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                false, false, Excel.XlSaveAsAccessMode.xlNoChange,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
        finally
        {
            oWB?.Close(false);
            Marshal.ReleaseComObject(oWB);
            oXL.Quit();
            Marshal.ReleaseComObject(oXL);
        }
    }
}
