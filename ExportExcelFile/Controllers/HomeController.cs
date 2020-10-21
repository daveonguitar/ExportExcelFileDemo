using System.Collections.Generic;
using Microsoft.AspNetCore.Mvc;
using ExportExcelFile.Models;
using System.IO;
using ClosedXML.Excel;

namespace ExportExcelFile.Controllers
{
    public class HomeController : Controller
    {
        private List<Employee> employees = new List<Employee>
        {
            new Employee {EmployeeId = 1, EmployeeName = "John", HireDate = "10-Jan-1997"},
            new Employee {EmployeeId = 2, EmployeeName = "Joe", HireDate = "23-Mar-1998"},
            new Employee {EmployeeId = 3, EmployeeName = "Steve", HireDate = "18-Jan-1999"},
            new Employee {EmployeeId = 4, EmployeeName = "Yancy", HireDate = "30-Apr-2001"},
            new Employee {EmployeeId = 5, EmployeeName = "Mukesh", HireDate = "01-Oct-2002"},
        };

        public IActionResult Index()
        {
            return Excel();
        }

        public IActionResult Excel()
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Employees");
                var currentRow = 1;
                worksheet.Cell(currentRow, 1).Value = "EmployeeId";
                worksheet.Cell(currentRow, 2).Value = "EmployeeName";
                worksheet.Cell(currentRow, 3).Value = "JoinDate";

                foreach (var employee in employees)
                {
                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = employee.EmployeeId;
                    worksheet.Cell(currentRow, 2).Value = employee.EmployeeName;
                    worksheet.Cell(currentRow, 3).Value = employee.HireDate;
                }

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();
                    return File(content,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "EmployeesList.xlsx");
                }
            }
        }
    }
}