using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Running;
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