using System.IO;
using HtmlAgilityPack;
using System.Collections.Generic;
using System.Globalization;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Security.Cryptography.X509Certificates;
using System.Diagnostics;

namespace CarChecker
{
    internal class Program
    {
        static void Main(string[] args)
        {

            List<Car> carList = GetCarsToList();

            string filePath = "C:\\Users\\karlo\\Desktop\\Car checker\\CarList.xlsx";

            ExcelWorksheet ws;
            ExcelPackage excel;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            if (!File.Exists(filePath))
            {
                excel = new ExcelPackage();
                Console.WriteLine("New xlsx file with name CarList has been created.");
            }
            else
            {
                excel = new ExcelPackage(filePath);
            }

            string today = DateTime.Now.ToString("ddMMyyyy");

            if (!excel.Workbook.Worksheets.Any(s => s.Name == today))
            {
                ws = excel.Workbook.Worksheets.Add(today);
                CreateHeaderRow(ws);
                Console.WriteLine($"New worksheet named {today} has been created.");
            }
            else
            {
                ws = excel.Workbook.Worksheets[today];
            }

            AddListDataToSheet(ws, carList);

            FileStream fileStream = File.Create(filePath);
            fileStream.Close();
            File.WriteAllBytes(filePath, excel.GetAsByteArray());

            var newCars = GetNewCars(excel);

            excel.Dispose();
            Console.WriteLine("All done.");
        }

        public static List<Car> GetCarsToList()
        {
            HtmlWeb web = new HtmlWeb();
            HtmlNode next;
            var carList = new List<Car>();
            var pageCounter = 1;

            do
            {
                HtmlDocument doc = web.Load("https://selection.renault.bg/" + "?page=" + pageCounter);
                var itemLinks = doc.DocumentNode.SelectNodes("//a[@class='itemLink']");
                next = doc.DocumentNode.SelectSingleNode("//a[@aria-label='Next']");

                if (itemLinks != null)
                {
                    foreach (var item in itemLinks)
                    {
                        var id = item.GetAttributeValue("href", "").Split("=", StringSplitOptions.RemoveEmptyEntries)[1].ToString().Trim();
                        var caption1 = item.ChildNodes.First(i => i.HasClass("caption1"));
                        var caption2 = item.ChildNodes.First(i => i.HasClass("caption2"));
                        var caption2a = item.ChildNodes.First(i => i.HasClass("caption2a")).InnerText.Split(" | ", StringSplitOptions.RemoveEmptyEntries);
                        var price = item.ChildNodes.First(i => i.HasClass("price"));
                        var url = item.GetAttributeValue("href", "").ToString().Trim();

                        carList.Add(new Car
                        {
                            Id = id,
                            Title = caption1.InnerText,
                            Description = caption2.InnerText,
                            AssemblyYear = caption2a[0],
                            Kilometers = int.Parse(caption2a[1].Replace("км", "").Trim().Replace(" ", "")),
                            FuelType = caption2a[2],
                            Price = double.Parse(price.InnerText.Replace("лв.", "").Trim().Replace(" ", "")),
                            Url = url
                        });
                    }
                }
                else
                {
                    Console.WriteLine($"Page {pageCounter} has been reached. End of pages.");
                    break;
                }

                pageCounter++;

            } while (next != null);

            Console.WriteLine($"List of {carList.Count} cars has been prepared.");

            return carList;
        }

        public static void AddListDataToSheet(ExcelWorksheet ws, List<Car> carList)
        {
            //Add data
            for (int t = 1; t <= carList.Count; t++)
            {
                ws.Cells[t + 1, 1].Value = carList[t-1].Id;
                ws.Cells[t + 1, 2].Value = carList[t-1].Title;
                ws.Cells[t + 1, 3].Value = carList[t-1].Description;
                ws.Cells[t + 1, 4].Value = carList[t-1].AssemblyYear;
                ws.Cells[t + 1, 5].Value = carList[t-1].Kilometers;
                ws.Cells[t + 1, 6].Value = carList[t-1].FuelType;
                ws.Cells[t + 1, 7].Value = carList[t-1].Price;
                ws.Cells[t + 1, 8].Value = carList[t-1].Url;
            }
            Console.WriteLine($"Cars data was successfuly written in sheet {ws.Name}.");
        }

        public static void CreateHeaderRow(ExcelWorksheet ws)
        {
            //Create excel file header row based on Car class public properties
            var props = typeof(Car).GetProperties();

            for (int p = 0; p < props.Length; p++)
            {
                ws.Cells[1, p + 1].Value = props[p].Name;
            }
        }

        public static Dictionary<string,string> GetNewCars(ExcelPackage excel)
        {
            Dictionary<string, string> newCars = new Dictionary<string, string>();

            var sheets = excel.Workbook.Worksheets.Select(ws => ws.Name).ToList();
            sheets.Sort();

            //Always keep two worksheets, before - after
            if (sheets.Count == 3)
            {
                var wsToDel = sheets[0];
                excel.Workbook.Worksheets.Delete(wsToDel);
            }

            var lastWs = excel.Workbook.Worksheets[sheets[2]];
            var prevWs = excel.Workbook.Worksheets[sheets[1]];

            for(var r = 1; r < lastWs.Rows.Count(); r++)
            {
                string id = lastWs.Cells[r, 1].Value.ToString();

                bool idExist = lastWs.SelectedRange["A:A"].Any(c => c.Value == id);

                if (!idExist)
                {
                    newCars.Add(lastWs.Cells[r, 2].Value.ToString(), lastWs.Cells[r, 8].Value.ToString());
                }
            }

            return newCars;
        }

        public class Car
        {
            public string Id { get; set; }
            public string Title { get; set; }
            public string Description { get; set; }
            public string AssemblyYear { get; set; }
            public int Kilometers { get; set; }
            public string FuelType { get; set; }
            public double Price { get; set; }
            public string Url { get; set; }
        }
    }
}
