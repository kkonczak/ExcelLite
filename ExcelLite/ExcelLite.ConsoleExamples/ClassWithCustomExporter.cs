namespace ExcelLite.ConsoleExamples
{
    public class ClassWithCustomExporter : ICustomCellExporter
    {
        public string TestText { get; set; }

        public string GetValue()
        {
            return TestText;
        }
    }
}
