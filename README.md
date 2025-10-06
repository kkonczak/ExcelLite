# 📦 ExcelLite – Lightweight Open Source Excel Generator for .NET

**ExcelLite** is an open-source C# library designed for **fast and memory-efficient generation of Excel (Open XML) files**. It focuses on **simplicity and performance**, making it ideal for scenarios where massive Excel files need to be created with minimal memory overhead.

## 🚀 Key Features

- ✅ **Minimalistic API** – Simple and intuitive interface for quick integration and easy use.
- 🧠 **Low RAM usage** – Designed to avoid keeping the entire document in memory.
- ⚡ **High performance** – Easily generates large-scale Excel files.
- 📄 **Open XML compliant** – Outputs standard `.xlsx` files.
- 📦 **No external dependencies** – Lightweight and dependency-free by default.
- 🌍 **Developed 100% in Europe** – Ensures full compliance with regional legal, security, and privacy requirements (e.g. GDPR).

## 💡 Why Use ExcelLite?

Most existing Excel libraries load the entire workbook into memory, which can be problematic when exporting large datasets. **ExcelLite** takes a streaming approach, writing XML directly to the output stream, allowing you to generate Excel files with **millions of rows** without overwhelming system resources.

This makes it perfect for:

- Generating large reports
- Exporting massive datasets from databases or logs
- Background/batch processing where performance and memory usage matter
- Compliance-conscious environments where software origin matters

## 🛠 Example Use Case
### Generate simple file
```csharp
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
```

### Generate 2 sheets in one file with freeze panes
```csharp
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
```

### Example attributes
```csharp
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
```
