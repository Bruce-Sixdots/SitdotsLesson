using Dapper;
using GemBox.Document;
using GemBox.Document.Tables;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.Sqlite;
using System.Data;

namespace GemBoxDemo.Controllers
{
    public class HomeController : Controller
    {
        private readonly string _dbPath = Path.Combine(Directory.GetCurrentDirectory(), "Data", "Demo.db");

        public IActionResult Index()
        {
            ViewBag.GeneratedFiles = GetGeneratedFiles();
            return View(new EventModel { Host = "王小明", Date = "2025/04/01" });
        }

        [HttpPost]
        public IActionResult Example1()
        {
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");
            var doc = new DocumentModel();
            var p = new Paragraph(doc, "test Write Something");
            doc.Sections.Add(new Section(doc, p));
            SaveDocument(doc, "test1");
            TempData["Message"] = "範例 1 完成!";
            return RedirectToAction("Index");
        }

        [HttpPost]
        public IActionResult Example2(EventModel model)
        {
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");
            var doc = new DocumentModel();
            var p = new Paragraph(doc,
                new Run(doc, "Dear All : "),
                new SpecialCharacter(doc, SpecialCharacterType.LineBreak),
                new Run(doc, $"本次的聚會將於 : {model.Date} 舉行"),
                new SpecialCharacter(doc, SpecialCharacterType.LineBreak),
                new Run(doc, $"主辦人為 : {model.Host}"),
                new SpecialCharacter(doc, SpecialCharacterType.LineBreak),
                new Run(doc, "請大家踴躍參加."),
                new SpecialCharacter(doc, SpecialCharacterType.LineBreak),
                new Run(doc, model.Host)
            );
            doc.Sections.Add(new Section(doc, p));
            SaveDocument(doc, "test2");
            TempData["Message"] = "範例 2 完成!";
            return RedirectToAction("Index");
        }

        [HttpPost]
        public IActionResult Example3()
        {
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");
            ComponentInfo.FreeLimitReached += (s, e) => e.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial;
            var doc = DocumentModel.Load(Path.Combine(Directory.GetCurrentDirectory(), "TemplateDoc", "BindingTable_Template.docx"));
            IEnumerable<dynamic> products;
            using (var connection = new SqliteConnection($"Data Source={_dbPath}"))
            {
                products = connection.Query<dynamic>("SELECT ProductId AS Productid, ProductName, UnitPrice, UnitsInStock FROM Products LIMIT 5");
                Console.WriteLine($"範例 3: 查詢到 {products.Count()} 筆資料");
                foreach (var p in products)
                {
                    Console.WriteLine($"Product: {p.Productid}, {p.ProductName}, {p.UnitPrice}, {p.UnitsInStock}");
                }
            }

            // 將 IEnumerable<dynamic> 轉換為 DataTable，並指定範圍名稱為 MyTable
            DataTable dt = new DataTable("MyTable");
            dt.Columns.Add("Productid", typeof(int));
            dt.Columns.Add("ProductName", typeof(string));
            dt.Columns.Add("UnitPrice", typeof(double));
            dt.Columns.Add("UnitsInStock", typeof(int));

            foreach (var product in products)
            {
                dt.Rows.Add(product.Productid, product.ProductName, product.UnitPrice, product.UnitsInStock);
            }

            doc.MailMerge.Execute(dt);
            doc.MailMerge.Execute(new { TotalAmount = "555" });
            SaveDocument(doc, "test3");
            TempData["Message"] = "範例 3 完成!";
            return RedirectToAction("Index");
        }
        [HttpPost]
        public IActionResult Example4()
        {
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");
            var doc = new DocumentModel();
            var pic = new Picture(doc, Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "HTML5.jpg"), 420, 256, LengthUnit.Pixel);
            var p = new Paragraph(doc, pic);
            doc.Sections.Add(new Section(doc, p));
            SaveDocument(doc, "test4");
            TempData["Message"] = "範例 4 完成!";
            return RedirectToAction("Index");
        }

        [HttpPost]
        public IActionResult Example5()
        {
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");
            var doc = new DocumentModel();
            var myTable = new Table(doc) { TableFormat = { PreferredWidth = new TableWidth(100, TableWidthUnit.Percentage) } };
            var header = new TableRow(doc);
            header.Cells.Add(new TableCell(doc, new Paragraph(doc, "Name")));
            header.Cells.Add(new TableCell(doc, new Paragraph(doc, "Type")));
            header.Cells.Add(new TableCell(doc, new Paragraph(doc, "Price")));
            myTable.Rows.Add(header);

            var products = GetProductArray();
            foreach (var item in products)
            {
                var row = new TableRow(doc);
                row.Cells.Add(new TableCell(doc, new Paragraph(doc, item.Name)));
                row.Cells.Add(new TableCell(doc, new Paragraph(doc, item.Type)));
                row.Cells.Add(new TableCell(doc, new Paragraph(doc, item.Price)));
                myTable.Rows.Add(row);
            }
            doc.Sections.Add(new Section(doc, myTable));
            SaveDocument(doc, "test5");  // 修正為 "test5"
            TempData["Message"] = "範例 5 完成!";
            return RedirectToAction("Index");
        }

        [HttpPost]
        public IActionResult Example6()
        {
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");
            ComponentInfo.FreeLimitReached += (s, e) => e.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial;
            var doc = DocumentModel.Load(Path.Combine(Directory.GetCurrentDirectory(), "TemplateDoc", "Resume_Template_Limit.docx"));
            var customer = new
            {
                Name = "王大呆aaaa",
                Sex = "男",
                Birthday = "2011/11/22",
                Email = "DaDai@uuu.com.tw",
                Address = "台灣aaaa",
                PhoneNumber = "0911222333",
                Education = "XX大學OO系",
                SalaryExpecte = "35000"
            };
            doc.MailMerge.Execute(customer);
            SaveDocument(doc, "test6");  // 修正為 "test6"
            TempData["Message"] = "範例 6 完成!";
            return RedirectToAction("Index");
        }

        private void SaveDocument(DocumentModel doc, string fileName)
        {
            var outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);
            doc.Save(Path.Combine(outputDir, $"{fileName}.docx"));
            doc.Save(Path.Combine(outputDir, $"{fileName}.pdf"));
        }

        private dynamic[] GetProductArray()
        {
            return new[]
            {
                new { Name = "HD", Type = "3TB", Price = "7500" },
                new { Name = "RAM", Type = "12G", Price = "3500" },
                new { Name = "CPU", Type = "i7 3200", Price = "8900" }
            };
        }

        private List<string> GetGeneratedFiles()
        {
            var outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            return Directory.Exists(outputDir)
                ? Directory.GetFiles(outputDir).Select(Path.GetFileName).OrderBy(f => f).ToList()
                : new List<string>();
        }
    }

    public class EventModel
    {
        public string Host { get; set; }
        public string Date { get; set; }
    }
}