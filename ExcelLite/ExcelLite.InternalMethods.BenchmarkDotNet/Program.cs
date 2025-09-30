using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Running;
using System.Text;

namespace ExcelLite.InternalMethods.BenchmarkDotNet
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var summary = BenchmarkRunner.Run<MethodBenchmarks>();
        }
    }

    [MemoryDiagnoser(true)]
    public class MethodBenchmarks
    {
        public MethodBenchmarks()
        {
        }

        //[Benchmark]
        public async Task StandardGetColumnNameBenchmark()
        {
            Enumerable.Repeat(1, 1000).Select(x => Enumerable.Range(0, 2000).Select(x => GetColumnName(x)).ToList()).ToList();
        }

        //[Benchmark]
        public async Task BetterGetColumnNameBenchmark()
        {
            Enumerable.Repeat(1, 1000).Select(x => Enumerable.Range(0, 2000).Select(x => BetterGetColumnName(x)).ToList()).ToList();
        }

        //[Benchmark]
        public async Task TestGetColumnNameBenchmark()
        {
            var hasDifferences = Enumerable.Repeat(1, 16380).Select(x => BetterGetColumnName(x) == GetColumnName(x)).ToList().Any(x => !x);

            if (hasDifferences)
            {
                throw new Exception("Bug in code!");
            }
        }

        public string GetColumnName(int columnIndex)
        {
            if (columnIndex < 0)
                throw new ArgumentOutOfRangeException(nameof(columnIndex), "Index must be non-negative.");

            string columnName = string.Empty;
            columnIndex++;

            while (columnIndex > 0)
            {
                int remainder = (columnIndex - 1) % 26;
                columnName = (char)(remainder + 'A') + columnName;
                columnIndex = (columnIndex - 1) / 26;
            }

            return columnName;
        }

        public string BetterGetColumnName(int columnIndex)
        {
            int remainder = 0;
            Span<char> columnNameChars = stackalloc char[3];
            int byteIndex = 2;
            columnIndex++;

            while (columnIndex > 0)
            {
                remainder = (columnIndex - 1) % 26;
                columnNameChars[byteIndex--] = (char)(remainder + 'A');
                columnIndex = (columnIndex - 1) / 26;
            }

            return new string(columnNameChars[(byteIndex + 1)..]);
        }

        [Benchmark]
        public void Write1()
        {
            var stream = new MemoryStream();
            StreamWriter streamWriter = new StreamWriter(stream);
            streamWriter.Write("asd");
            streamWriter.Write("asd");
            streamWriter.Write("asd");
            streamWriter.Write("asd");
            streamWriter.Write("asd");
            streamWriter.Write("asd");
            streamWriter.Write("asd");
            streamWriter.Write("asd");
            streamWriter.Write("asd");
            streamWriter.Write("asd");
            streamWriter.Write("asd");
        }

        [Benchmark]
        public void Write2()
        {
            var stream = new MemoryStream();
            StreamWriter streamWriter = new StreamWriter(stream);
            StringBuilder builder = new StringBuilder();
            builder.Append("asd");
            builder.Append("asd");
            builder.Append("asd");
            builder.Append("asd");
            builder.Append("asd");
            builder.Append("asd");
            builder.Append("asd");
            builder.Append("asd");
            builder.Append("asd");
            builder.Append("asd");
            builder.Append("asd");
            streamWriter.Write("asd");
        }
    }
}