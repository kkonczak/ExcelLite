using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Diagnostics;

namespace ExcelLite.ConsoleExamples
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            // await GenerateSheet() // Example 1
            await GenerateSheetWithFreezePanes(); // Example 2
            // await MultipleSheetsTest();  // example 3
            // await SheetFromDbTest(); //example 4
            // await SheetManyRowsTest(); //example 5
            // await SheetWithCustomClassInData(); // example 6
            //await SheetWithNestedClasses(); // example 7
            // await ManyRecordsTest();
            // await RecordTest();
            //await AsyncEnumerableTest();
        }

        private static async Task GenerateSheet()
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

            await ExcelLite.Export("test.xlsx", data);
        }

        private static async Task GenerateSheetWithFreezePanes()
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

            var sheet1 = new Sheet("Arkusz1", data);
            sheet1.View.FreezePanes.YSplit = 3;

            await ExcelLite.Export("test.xlsx", new Workbook(new Sheet[] { sheet1 }));
        }

        private static async Task MultipleSheetsTest()
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
                    DateOnly = new DateOnly(2025,07,01),
                    PropertyWithExporter = new ClassWithCustomExporter()
                    {
                        TestText = "Content of TestText"
                    }
                },
            };

            var data2 = new List<SalaryTestClass>()
            {
                new SalaryTestClass
                {
                    FirstName = "Jan",
                    LastName = "Nowak",
                    Income = 6000
                },
                new SalaryTestClass
                {
                    FirstName = "Jan",
                    LastName = "Kowalski",
                    Income = 6100
                },
                new SalaryTestClass
                {
                    FirstName = "Jan",
                    LastName = "Adamczewski",
                    Income = 6200
                }
            };

            var sheet1 = new Sheet("Arkusz1", data);
            sheet1.View.FreezePanes.YSplit = 3;

            await ExcelLite.Export("test.xlsx", new Workbook(new Sheet[] { new Sheet("Salary", data2), sheet1 }));
        }

        private static async Task SheetFromDbTest()
        {
            var _contextOptions = new DbContextOptionsBuilder<SalaryDbContext>()
                .UseInMemoryDatabase("SalaryDbTest")
                .ConfigureWarnings(b => b.Ignore(InMemoryEventId.TransactionIgnoredWarning))
            .Options;

            using var context = new SalaryDbContext(_contextOptions);

            context.Database.EnsureDeleted();
            context.Database.EnsureCreated();

            context.Salaries.AddRange(new DbSalary
            {
                FirstName = "John",
                LastName = "Doe",
                Salary = 123.4m
            },
            new DbSalary
            {
                FirstName = "Johnny",
                LastName = "Jobs",
                Salary = 9230.4m
            });

            context.SaveChanges();

            await ExcelLite.Export("test.xlsx", context.Salaries.Where(x => x.Salary > 500m));
        }

        public static async Task SheetManyRowsTest()
        {
            await ExcelLite.Export("test.xlsx", Enumerable.Range(0, 12345).Select(x => new { Column1 = x, Column2 = x, Column3 = x }));
        }

        public static async Task SheetWithCustomClassInData()
        {
            var data = new List<ClassForSheetWithCustomClassInData>()
            {
                new ClassForSheetWithCustomClassInData()
                {
                    ClassAField1 = "field1",
                    NestedClass = new NestedClassForSheetWithCustomClassInData()
                    {
                        ClassBField1 = "NestedClass->ClassBField1",
                        ClassBField2 = "NestedClass->ClassBField2",
                        ClassC = new ClassC()
                        {
                            ClassCField1 = "NestedClass->ClassC->ClassCField1",
                            ClassCField2 = "NestedClass->ClassC->ClassCField2",
                            ClassCField3 = "NestedClass->ClassC->ClassCField3"
                        }
                    },
                    ClassAField2 = "field2"
                }
            };

            await ExcelLite.Export("test.xlsx", data);
        }

        public static async Task SheetWithNestedClasses()
        {
            var data = new List<RecordWithNestedClasses>()
            {
                new RecordWithNestedClasses
                {
                    PersonalData = new PersonalData
                    {
                        FirstName = "Jan",
                        LastName = "Nowak",
                        PhoneNumber = "333444555"
                    },
                    Address = new Address
                    {
                        BuildingNumber = "1",
                        City = "Warsaw",
                        Country = "Poland",
                        PostalCode = "00-000",
                        StreetName = "ABC"
                    }
                },new RecordWithNestedClasses
                {
                    PersonalData = new PersonalData
                    {
                        FirstName = "Jan",
                        LastName = "Kowalski",
                        PhoneNumber = "133444555"
                    },
                    Address = new Address
                    {
                        BuildingNumber = "1",
                        City = "Łódź",
                        Country = "Poland",
                        PostalCode = "00-000",
                        StreetName = "ABC"
                    }
                },
                new RecordWithNestedClasses
                {
                    PersonalData = new PersonalData
                    {
                        FirstName = "John",
                        LastName = "Doe",
                        PhoneNumber = "2435368493"
                    },
                    Address = new Address
                    {
                        BuildingNumber = "1",
                        City = "New York",
                        Country = "USA",
                        PostalCode = "23456",
                        StreetName = "ABC"
                    }
                }
            };

            //await ExcelLite.Export("test.xlsx", data);
            await ExcelLite.Export("test.xlsx", new Workbook(
                new List<Sheet>{
                    new Sheet("test", data)
                    {
                        UseBorders = true
                    }
                })
            );
        }

        public static async Task ManyRecordsTest()
        {
            var data = Enumerable.Range(0, 980000).Select(x => new { Value = x, Vaue2 = x, Value3 = "testtest" });
            await ExcelLite.Export("test.xlsx", data);
        }

        public static async Task RecordTest()
        {
            var data = Enumerable.Range(0, 100).Select(x => new RecordT("abc", "def", x));
            await ExcelLite.Export("test.xlsx", data);
        }

        public static async Task AsyncEnumerableTest()
        {
            await ExcelLite.Export("test.xlsx", TestAsyncEnumerable());
        }

        private static async IAsyncEnumerable<RecordT> TestAsyncEnumerable()
        {
            for (int i = 1; i <= 60; i++)
            {
                await Task.Delay(10);
                yield return new RecordT("abc", "def", i); 
            }
        }
    }

    public record RecordT(string Field1, string Field2, int Field3) { }
}