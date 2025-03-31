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
            // �N�۹���|�ഫ��������|
            string relativePath = _dbPath.Replace("Data Source=", "");
            _dbPath = "Data Source=" + Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), relativePath));
        }

        // �d�� 1�G�q��Ʈw�d�߸�ƨüg�J Excel
        [HttpPost]
        public IActionResult Example1()
        {
            try
            {
                SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
                SpreadsheetInfo.FreeLimitReached += (s, e) => e.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial;

                // �d�߸�Ʈw
                IEnumerable<dynamic> products;
                using (var connection = new SqliteConnection(_dbPath))
                {
                    connection.Open();
                    products = connection.Query<dynamic>("SELECT ProductId, ProductName, UnitPrice, UnitsInStock FROM Products LIMIT 10");

                    Console.WriteLine($"�d�� 1: �d�ߨ� {products.Count()} �����");
                    foreach (var p in products)
                    {
                        Console.WriteLine($"Product: {p.ProductId}, {p.ProductName}, {p.UnitPrice}, {p.UnitsInStock}");
                    }
                }

                if (products == null || !products.Any())
                {
                    TempData["Message"] = "�d�L���";
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
                Console.WriteLine($"�ɮױN�x�s��: {filePath}");
                xlsx.Save(filePath);

                TempData["Message"] = "Excel �d�� 1 �����I";
                return RedirectToAction("Index", "Home");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"���~: {ex.Message}");
                TempData["Message"] = $"Excel �d�� 1 ����: {ex.Message}";
                return RedirectToAction("Index", "Home");
            }
        }

        // �d�� 2�GŪ�� Excel �ɮר���ܸ��
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
                TempData["Message"] = $"Excel �d�� 2 ����: {ex.Message}";
                return RedirectToAction("Index", "Home");
            }
        }

        // �d�� 3�G�q Excel ������ƨ���ܦb GridView
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
                TempData["Message"] = $"Excel �d�� 3 ����: {ex.Message}";
                return RedirectToAction("Index", "Home");
            }
        }

        // �d�� 4�G�ʺA�ͦ� Excel �ɮ�
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

                TempData["Message"] = "Excel �d�� 4 �����I";
                return RedirectToAction("Index", "Home");
            }
            catch (Exception ex)
            {
                TempData["Message"] = $"Excel �d�� 4 ����: {ex.Message}";
                return RedirectToAction("Index", "Home");
            }
        }
    }
}