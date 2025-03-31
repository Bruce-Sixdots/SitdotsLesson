using Microsoft.Data.Sqlite;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddControllersWithViews();

var app = builder.Build();

// 初始化 SQLite 資料庫
var dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
Directory.CreateDirectory(dataDir);
var dbPath = Path.Combine(dataDir, "Demo.db");
using (var connection = new SqliteConnection($"Data Source={dbPath}"))
{
    connection.Open();
    var command = connection.CreateCommand();
    command.CommandText = @"
        CREATE TABLE IF NOT EXISTS Products (
            ProductId INTEGER PRIMARY KEY,
            ProductName TEXT NOT NULL,
            UnitPrice REAL NOT NULL,
            UnitsInStock INTEGER NOT NULL
        );
        INSERT OR IGNORE INTO Products (ProductId, ProductName, UnitPrice, UnitsInStock) VALUES
            (1, 'Product A', 10.5, 100),
            (2, 'Product B', 20.0, 50),
            (3, 'Product C', 15.0, 75),
            (4, 'Product D', 25.0, 30),
            (5, 'Product E', 30.0, 20);
    ";
    command.ExecuteNonQuery();
}


// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles();

app.UseRouting();

app.UseAuthorization();

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Home}/{action=Index}/{id?}");

app.Run();
