using System.IO;
using HtmlAgilityPack;
using System.Collections.Generic;
using System.Globalization;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace CarChecker
{
    internal class Program
    {
        static void Main(string[] args)
        {
            HtmlWeb web = new HtmlWeb();
            var carList = new List<Car>();
            HtmlNode next = null;
            var pageCounter = 1;

            do
            {
                HtmlDocument doc = web.Load("https://selection.renault.bg/" + "?page=" + pageCounter);
                var itemLinks = doc.DocumentNode.SelectNodes("//a[@class='itemLink']");
                next = doc.DocumentNode.SelectSingleNode("//a[@aria-label='Next']");

                if(itemLinks != null)
                {
                    foreach (var item in itemLinks)
                    {
                        var id = item.GetAttributeValue("href", "").Split("=", StringSplitOptions.RemoveEmptyEntries)[1].ToString().Trim();

                        var caption1 = item.ChildNodes.First(i => i.HasClass("caption1"));
                        var caption2 = item.ChildNodes.First(i => i.HasClass("caption2"));
                        var caption2a = item.ChildNodes.First(i => i.HasClass("caption2a")).InnerText.Split(" | ", StringSplitOptions.RemoveEmptyEntries);
                        var price = item.ChildNodes.First(i => i.HasClass("price"));

                        carList.Add(new Car
                        {
                            Id = id,
                            Title = caption1.InnerText,
                            Description = caption2.InnerText,
                            AssemblyYear = caption2a[0],
                            Kilometers = caption2a[1],
                            FuelType = caption2a[2],
                            Price = price.InnerText
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

            Console.WriteLine($"List of {carList.Count} cars was prepared.");

            string filePath = "C:\\Users\\karlo\\Desktop\\Car checker\\CarList.xlsx";

            //TODO: Continue from here
            if (File.Exists(filePath))
            {

            }
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage excel = new ExcelPackage();


            var ws = excel.Workbook.Worksheets.Add(DateTime.Now.ToString("MM/dd/yyyy HH:mm"));

            //Create excel file header row based on Car class public properties
            var props = typeof(Car).GetProperties();

            for (int p = 0; p < props.Length - 1; p++)
            {
                ws.Cells[1, p + 1].Value = props[p].Name;
            }

            //Add data
            for (int t = 0; t < carList.Count - 1; t++)
            {
                ws.Cells[t + 1, 1].Value = carList[t].Id;
                ws.Cells[t + 1, 2].Value = carList[t].Title;
                ws.Cells[t + 1, 3].Value = carList[t].Description;
                ws.Cells[t + 1, 4].Value = carList[t].AssemblyYear;
                ws.Cells[t + 1, 5].Value = carList[t].Kilometers;
                ws.Cells[t + 1, 6].Value = carList[t].FuelType;
                ws.Cells[t + 1, 7].Value = carList[t].Price;
            }

            //This goes up
            if (File.Exists(filePath))
            {
                excel = new ExcelPackage(filePath);

            }
            else
            {
                FileStream fileStream = File.Create(filePath);
                fileStream.Close();

                File.WriteAllBytes(filePath, excel.GetAsByteArray());
            }

            excel.Dispose();

        }

        public class Car
        {
            public string Id { get; set; }
            public string Title { get; set; }
            public string Description { get; set; }
            public string AssemblyYear { get; set; }
            public string Kilometers { get; set; }
            public string FuelType { get; set; }
            public string Price { get; set; }
        }
    }
}
