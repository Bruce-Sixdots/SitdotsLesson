using GemBox.Spreadsheet;
using Microsoft.AspNetCore.Mvc;
using System.Data;
using Dapper;
using Microsoft.Data.Sqlite;

namespace GemBoxDemo.Controllers
{
    public class ExcelController : Controller
    {
        private readonly string _dbPath;

        public ExcelController(IConfiguration configuration)
        {
            _dbPath = configuration.GetConnectionString("SQLiteConnection");
            // 將相對路徑轉換為絕對路徑
            string relativePath = _dbPath.Replace("Data Source=", "");
            _dbPath = "Data Source=" + Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), relativePath));
        }

        // 範例 1：從資料庫查詢資料並寫入 Excel
        [HttpPost]
        public IActionResult Example1()
        {
            try
            {
                SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
                SpreadsheetInfo.FreeLimitReached += (s, e) => e.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial;

                // 查詢資料庫
                IEnumerable<dynamic> products;
                using (var connection = new SqliteConnection(_dbPath))
                {
                    connection.Open();
                    products = connection.Query<dynamic>("SELECT ProductId, ProductName, UnitPrice, UnitsInStock FROM Products LIMIT 10");

                    Console.WriteLine($"範例 1: 查詢到 {products.Count()} 筆資料");
                    foreach (var p in products)
                    {
                        Console.WriteLine($"Product: {p.ProductId}, {p.ProductName}, {p.UnitPrice}, {p.UnitsInStock}");
                    }
                }

                if (products == null || !products.Any())
                {
                    TempData["Message"] = "查無資料";
                    return RedirectToAction("Index", "Home");
                }

                DataTable dt = new DataTable("MyTable");
                dt.Columns.Add("ProductId", typeof(int));
                dt.Columns.Add("ProductName", typeof(string));
                dt.Columns.Add("UnitPrice", typeof(double));
                dt.Columns.Add("UnitsInStock", typeof(int));

                foreach (var product in products)
                {
                    dt.Rows.Add(product.ProductId, product.ProductName, product.UnitPrice, product.UnitsInStock);
                }

                ExcelFile xlsx = new ExcelFile();
                ExcelWorksheet mySheet = xlsx.Worksheets.Add("sheet1");
                mySheet.InsertDataTable(dt, new InsertDataTableOptions
                {
                    StartColumn = 2,
                    StartRow = 2,
                    ColumnHeaders = true
                });

                string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "OutputExcel");
                Directory.CreateDirectory(outputDir);
                string filePath = Path.Combine(outputDir, "test3.xlsx");
                Console.WriteLine($"檔案將儲存到: {filePath}");
                xlsx.Save(filePath);

                TempData["Message"] = "Excel 範例 1 完成！";
                return RedirectToAction("Index", "Home");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"錯誤: {ex.Message}");
                TempData["Message"] = $"Excel 範例 1 失敗: {ex.Message}";
                return RedirectToAction("Index", "Home");
            }
        }

        // 範例 2：讀取 Excel 檔案並顯示資料
        [HttpPost]
        public IActionResult Example2()
        {
            try
            {
                SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
                SpreadsheetInfo.FreeLimitReached += (s, e) => e.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial;

                string filePath = Path.Combine(Directory.GetCurrentDirectory(), "OutputExcel", "test1.xlsx");
                ExcelFile xlsx = ExcelFile.Load(filePath);
                ExcelWorksheet mySheet = xlsx.Worksheets["sheet1"];

                List<string> results = new List<string>();
                foreach (ExcelRow item in mySheet.Rows)
                {
                    if (item.Cells[0].Value != null && item.Cells[1].Value != null)
                    {
                        string str = $"{item.Cells[0].Value}, {item.Cells[1].Value}";
                        results.Add(str);
                    }
                }

                string b1Value = mySheet.Cells["B1"].Value?.ToString() ?? "N/A";
                ViewBag.Results = results;
                ViewBag.B1Value = b1Value;
                return View("DisplayExcelData");
            }
            catch (Exception ex)
            {
                TempData["Message"] = $"Excel 範例 2 失敗: {ex.Message}";
                return RedirectToAction("Index", "Home");
            }
        }

        // 範例 3：從 Excel 提取資料並顯示在 GridView
        [HttpPost]
        public IActionResult Example3()
        {
            try
            {
                SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
                SpreadsheetInfo.FreeLimitReached += (s, e) => e.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial;

                string filePath = Path.Combine(Directory.GetCurrentDirectory(), "OutputExcel", "test3.xlsx");
                ExcelFile xlsx = ExcelFile.Load(filePath);
                ExcelWorksheet mySheet = xlsx.Worksheets["sheet1"];

                DataTable dt = new DataTable();
                dt.Columns.Add("ProductId", typeof(string));
                dt.Columns.Add("ProductName", typeof(string));
                dt.Columns.Add("UnitPrice", typeof(double));
                dt.Columns.Add("UnitsInStock", typeof(int));

                ExtractToDataTableOptions options = new ExtractToDataTableOptions(3, 2, 10);
                mySheet.ExtractToDataTable(dt, options);

                return View("DisplayGridView", dt);
            }
            catch (Exception ex)
            {
                TempData["Message"] = $"Excel 範例 3 失敗: {ex.Message}";
                return RedirectToAction("Index", "Home");
            }
        }

        // 範例 4：動態生成 Excel 檔案
        [HttpPost]
        public IActionResult Example4()
        {
            try
            {
                SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
                SpreadsheetInfo.FreeLimitReached += (s, e) => e.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial;

                ExcelFile xlsx = new ExcelFile();
                ExcelWorksheet mySheet = xlsx.Worksheets.Add("sheet1");

                mySheet.Cells[0, 0].Value = "ProductName";
                mySheet.Cells[0, 0].Style.FillPattern.SetSolid(SpreadsheetColor.FromName(ColorName.Orange));
                mySheet.Cells[0, 1].Value = "Price";
                mySheet.Cells[0, 1].Style.FillPattern.SetSolid(SpreadsheetColor.FromName(ColorName.Orange));

                Random rnd = new Random();
                for (int i = 1; i < 20; i++)
                {
                    mySheet.Cells[i, 0].Value = "Product" + i.ToString();
                    mySheet.Cells[i, 1].Value = rnd.Next(1, 1000);
                }

                string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "OutputExcel");
                Directory.CreateDirectory(outputDir);
                string filePath = Path.Combine(outputDir, "test1.xlsx");
                xlsx.Save(filePath);

                TempData["Message"] = "Excel 範例 4 完成！";
                return RedirectToAction("Index", "Home");
            }
            catch (Exception ex)
            {
                TempData["Message"] = $"Excel 範例 4 失敗: {ex.Message}";
                return RedirectToAction("Index", "Home");
            }
        }
    }
}