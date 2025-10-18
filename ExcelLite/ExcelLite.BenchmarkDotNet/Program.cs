using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Running;
using ClosedXML.Excel;
using MiniExcelLibs;

namespace ExcelLite.BenchmarkDotNet
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var summary = BenchmarkRunner.Run<ExcelBenchmarks>();
        }
    }

    [MemoryDiagnoser(true)]
    public class ExcelBenchmarks
    {
        private IEnumerable<TestClass> _data;

        public ExcelBenchmarks()
        {
            _data = Enumerable.Range(0, 1000000)
                .Select(x => new TestClass
                {
                    Age = 20,
                    Id = x,
                    FirstName = "Joe",
                    LastName = "Doe",
                    Column5 = "Hello World",
                    Column6 = "Hello World",
                    Column7 = "Hello World",
                    Column8 = "Hello World",
                    Column9 = "Hello World",
                    Column10 = "Hello World",
                    Column11 = "Hello World",
                    Column12 = "Hello World",
                });
        }

        [Benchmark]
        public async Task ExcelLiteBenchmark()
        {
            var stream = Stream.Null;
            await ExcelLite.Export(stream, _data);
        }

        [Benchmark]
        public async Task MiniExcelBenchmark()
        {
            var stream = Stream.Null;
            await MiniExcel.SaveAsAsync(stream, _data);
        }

        //[Benchmark]
        public async Task ClosedXMLBenchmark()
        {
            var stream = Stream.Null;

            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            int rowI = 2;
            foreach (var row in _data)
            {
                var cell = ws.Cell("A" + rowI);
                cell.Value = row.Id;
                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;

                cell = ws.Cell("B" + rowI);
                cell.Value = row.Age;
                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;

                cell = ws.Cell("C" + rowI);
                cell.Value = row.FirstName;
                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;

                cell = ws.Cell("D" + rowI);
                cell.Value = row.LastName;
                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;

                cell = ws.Cell("E" + rowI);
                cell.Value = row.Column5;
                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;

                cell = ws.Cell("F" + rowI);
                cell.Value = row.Column6;
                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;

                cell = ws.Cell("G" + rowI);
                cell.Value = row.Column7;
                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;

                cell = ws.Cell("H" + rowI);
                cell.Value = row.Column8;
                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;

                cell = ws.Cell("I" + rowI);
                cell.Value = row.Column9;
                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;

                cell = ws.Cell("J" + rowI);
                cell.Value = row.Column10;
                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;

                cell = ws.Cell("K" + rowI);
                cell.Value = row.Column11;
                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;

                cell = ws.Cell("L" + rowI);
                cell.Value = row.Column12;
                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;

                rowI++;
            }

            wb.SaveAs(stream);
        }
    }

    public class TestClass
    {
        public int Id { get; set; }

        public int Age { get; set; }

        public string? FirstName { get; set; }

        public string? LastName { get; set; }

        public string? Column5 { get; set; }

        public string? Column6 { get; set; }

        public string? Column7 { get; set; }

        public string? Column8 { get; set; }

        public string? Column9 { get; set; }

        public string? Column10 { get; set; }

        public string? Column11 { get; set; }

        public string? Column12 { get; set; }
    }
}