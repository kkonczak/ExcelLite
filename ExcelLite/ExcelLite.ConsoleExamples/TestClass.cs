using ExcelLite.Attributes;

namespace ExcelLite.ConsoleExamples
{
    public class TestClass
    {
        [GroupColumnName("Personal data", 0)]
        [GroupColumnName("First and Last Name", 1)]
        public string FirstName { get; set; }

        [GroupColumnName("Personal data", 0)]
        [GroupColumnName("First and Last Name", 1)]
        public string LastName { get; set; }

        [GroupColumnName("Personal data", 0)]
        public int Income { get; set; }


        [ColumnName("CustomName")]
        [ColumnFormat("00.00")]
        public float Float { get; set; }

        //[ColumnPosition(2)]
        public bool OtherBoolean { get; set; }

        [GroupColumnName("Address", 1)]
        public string Street { get; set; }

        [GroupColumnName("Address", 1)]
        public string City { get; set; }

        [GroupColumnName("Address", 1)]
        public string PostalCode { get; set; }

        [GroupColumnName("Address", 1)]
        public string BuildingNumber { get; set; }

        public DateTime DateTime { get; set; }

        public DateOnly DateOnly { get; set; }

        public TimeOnly TimeOnly { get; set; }
    }
}
