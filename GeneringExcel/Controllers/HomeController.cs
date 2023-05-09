using System.Data;
using System.Diagnostics;
using ClosedXML.Excel;
using Entities;
using GeneringExcel.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace GeneringExcel.Controllers;

public class HomeController : Controller
{
    private readonly ILogger<HomeController> _logger;
    private readonly ApplicationDbContext _context;

    public HomeController(ILogger<HomeController> logger, ApplicationDbContext context)
    {
        _logger = logger;
        _context = context;
    }

    public IActionResult Index()
    {
        return View();
    }

    [HttpGet]
    public async Task<FileResult> ExportPersonsToExcel(){
        var persons = await _context.Persons.ToListAsync();
        var nameFile = "Persons.xlsx";
        return GenerateExcel(nameFile,persons);
    }

    [HttpPost]
    public async Task<IActionResult> ImportPersonsFromExcel(IFormFile excel){
        
        using(var workbook = new XLWorkbook(excel.OpenReadStream()))
        {
            var sheet = workbook.Worksheet(1);

            var firstRowUsed = sheet.FirstRowUsed().RangeAddress.FirstAddress.RowNumber;
            var lastRowUsed = sheet.LastRowUsed().RangeAddress.FirstAddress.RowNumber;

            var persons = new List<Person>();

            for(var i = firstRowUsed+1; i <= lastRowUsed; i++){
                var row = sheet.Row(i);
                var person = new Person();

                person.Name = row.Cell(1).GetString();
                person.Salary = row.Cell(2).GetValue<decimal>();
                person.DateOfBirth = row.Cell(3).GetDateTime();
                persons.Add(person);
            }

            _context.AddRange(persons);
            await _context.SaveChangesAsync();
        }

        return View("Index");
        
    }

    private FileResult GenerateExcel(string nameFile, IEnumerable<Person> persons)
    {
        DataTable dataTable = new DataTable("Persons");
        dataTable.Columns.AddRange(new DataColumn[]{
            new DataColumn("Id"),
            new DataColumn("Name")
        });

        foreach (var person in persons)
        {
            dataTable.Rows.Add(person.Id, person.Name);
        }

        using (XLWorkbook wb = new XLWorkbook())
        {
            wb.Worksheets.Add(dataTable);

            using (MemoryStream stream = new MemoryStream())
            {
                wb.SaveAs(stream);
                return File(stream.ToArray(),
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    nameFile);
            }
        }
    }

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
