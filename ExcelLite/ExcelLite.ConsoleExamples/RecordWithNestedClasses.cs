namespace ExcelLite.ConsoleExamples
{
    public class RecordWithNestedClasses
    {
        public PersonalData PersonalData { get; set; }
        public Address Address { get; set; }
    }

    public class PersonalData
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string PhoneNumber { get; set; }
    }

    public class Address
    {
        public string StreetName { get; set; }
        public string BuildingNumber { get; set; }
        public string PostalCode { get; set; }
        public string City { get; set; }
        public string Country { get; set; }
    }
}
