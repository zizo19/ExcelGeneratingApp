using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using ExcelGeneratingApp.Models;

namespace ExcelGeneratingApp.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            List<Employee> employees = new List<Employee>();
            employees.Add(new Employee() { ID = 100 ,Name="TOUFIK", Age=32,Salary=12000,Department="INFO"}) ;
            employees.Add(new Employee() { ID = 200 ,Name="MERYEM", Age=24,Salary=10000,Department="CIVIL"}) ;
            employees.Add(new Employee() { ID = 300 ,Name="FERDAWS", Age=20,Salary=11000,Department="INDUS"}) ;
            employees.Add(new Employee() { ID = 400 ,Name="MOHAMED ALI", Age=21,Salary=9000,Department="INFO"}) ;
            ExcelLib excel = new ExcelLib();

            excel.Generate(employees.Cast<object>().ToList(), "List of employees");
            
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}