using ExcelLite.Attributes;

namespace ExcelLite.ConsoleExamples
{
    public class SalaryTestClass
    {
        public string FirstName { get; set; }

        public string LastName { get; set; }

        public int Income { get; set; }

        [ColumnIgnore]
        public int? HiddenField { get; set; }
    }
}
