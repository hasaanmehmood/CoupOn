using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using Test_Backend.Models;
using Newtonsoft.Json.Linq;
using System;
using System.IO;
using Newtonsoft.Json;
using OfficeOpenXml;
using Newtonsoft.Json;
using System.Data;
using System.Linq;
using OfficeOpenXml;
using System.Data;
using System.Linq;

public class ExcelConverter
{
    public DataTable ConvertToDataTable(string filePath)
    {
        // Create a new DataTable to store the data from the excel file
        DataTable dataTable = new DataTable();
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


        // Open the excel file using EPPlus
        using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
        {
            // Get the first worksheet in the excel file
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
           // var worksheet = package.Workbook.Worksheets.FirstOrDefault();
            if (worksheet != null && worksheet.Name == "Shop")
            {
               
                // Loop through the rows of the worksheet
                for (int row = 1; row <= worksheet.Dimension.Rows; row++)
                {
                     if (row == 1)
                    {
                        // This is the first row, so we need to create the column names
                        for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                        {
                            dataTable.Columns.Add(worksheet.Cells[row, col].Value.ToString());
                        }
                    }
                    else
                    {
                        // This is not the first row, so we need to add the data to the DataTable
                        DataRow dataRow = dataTable.NewRow();
                        for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                        {
                            dataRow[col - 1] = worksheet.Cells[row, col].Value;
                        }
                        dataTable.Rows.Add(dataRow);
                    }
                }
            }

            return dataTable;

        }

    }

}
public class Product
{
    public string StartDate { get; set; }
    public string EndDate { get; set; }
    public string ImageUrl { get; set; }
}
namespace Test_Backend.Controllers
{
    

    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }
        public IActionResult Index()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string filePath = @"C:\Users\hasee\source\repos\Test Backend\Test Backend\test.xlsx";
            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                Console.WriteLine("Sheet name: " + worksheet.Name);
                Console.WriteLine("Number of sheets: " + package.Workbook.Worksheets.Count);
                // ...
            }

            
            var excelConverter = new ExcelConverter();
            DataTable dataTable = excelConverter.ConvertToDataTable(filePath);

            var json = JsonConvert.SerializeObject(dataTable);
            return Content(json);
        }

        public IActionResult Fun()
        {
            List<Product> products = new List<Product>();
            string filePath = @"C:\Users\hasee\source\repos\Test\Test\test.xlsx";
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];

                for (int i = 2; i <= worksheet.Dimension.Rows; i++)
                {
                    products.Add(new Product
                    {
                        StartDate = worksheet.Cells[i, 1].Value.ToString(),
                        EndDate = worksheet.Cells[i, 2].Value.ToString(),
                        ImageUrl = worksheet.Cells[i, 3].Value.ToString()
                    });
                }
            }

            //Return the data to the view
            return View(products);
        }

        //public IActionResult Index()
        //{
        //    return View();
        //}

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}