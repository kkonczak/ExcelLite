using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Diagnostics;
using System.Reflection.Metadata;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace ExcelLite.ConsoleExamples
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            // await GenerateSheet() // Example 1
            // await GenerateSheetWithFreezePanes() // Example 2
            await MultipleSheetsTest();  // example 3
            //await SheetFromDbTest(); //example 4
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
                    DateOnly = new DateOnly(2025,07,01)
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

            await ExcelLite.Export("test.xlsx", new Workbook(new Sheet[] { new Sheet("Salary", data2), sheet1  }));
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
    }
}