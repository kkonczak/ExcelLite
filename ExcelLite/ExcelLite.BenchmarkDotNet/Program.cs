using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Running;
using MiniExcelLibs;
using System.IO;

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
            _data = Enumerable.Range(0, 300000)
                .Select(x => new TestClass
                {
                    Age = 20,
                    Id = x,
                    FirstName = "Joe",
                    LastName = "Doe"
                });
        }
        [Benchmark]
        public async Task ExcelLiteBenchmark()
        {
            var stream = new MemoryStream();
            await ExcelLite.Export(stream, _data);
        }

        [Benchmark]
        public async Task MiniExcelBenchmark()
        {
            var stream = new MemoryStream();
            await MiniExcel.SaveAsAsync(stream, _data);
        }
    }

    public class TestClass
    {
        public int Id { get; set; }

        public int Age { get; set; }

        public string? FirstName { get; set; }

        public string? LastName { get; set; }
    }
}