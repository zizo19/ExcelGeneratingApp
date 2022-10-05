using System.ComponentModel;
namespace ExcelGeneratingApp.Models
{
    public class Employee
    {
        [DisplayName("Identity number")]
        public int ID { get; set; }
        [DisplayName("Full name")]
        public string Name { get; set; }
        [DisplayName("Age")]
        public int Age { get; set; }
        [DisplayName("Salary")]
        public float Salary { get; set; }
        [DisplayName("Department name")]
        public string Department { get; set; }
        
    }
}