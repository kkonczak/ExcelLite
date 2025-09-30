namespace ExcelLite.ConsoleExamples
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            var data = new List<TestClass>()
            {
                new TestClass
                {
                    FirstName = "Jan",
                    LastName = "Nowak",
                    Income = 6000,
                    Float = 34.232f,
                    Street="asdasdasdasdasdas"
                },
                new TestClass
                {
                    FirstName = "Jan",
                    LastName = "Kowalski",
                    Income = 7000,
                    Float = 34.232f,
                    OtherBoolean = true,
                    DateTime= DateTime.Now,
                    TimeOnly = new TimeOnly(21,37,00),
                    DateOnly = new DateOnly(2025,07,01)
                },
            };

            // await ExcelLite.Export("test.xlsx", data); // Example 1
            var sheet1 = new Sheet("Arkusz1", data);
            sheet1.View.FreezePanes.YSplit = 3;
            await ExcelLite.Export("test.xlsx", new Workbook(new Sheet[] { sheet1 }));
        }
    }
}