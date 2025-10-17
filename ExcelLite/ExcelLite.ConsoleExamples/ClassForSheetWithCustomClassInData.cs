namespace ExcelLite.ConsoleExamples
{
    public class ClassForSheetWithCustomClassInData
    {
        public string ClassAField1 { get; set; }

        public NestedClassForSheetWithCustomClassInData NestedClass { get; set; }

        public string ClassAField2 { get; set; }
    }

    public class NestedClassForSheetWithCustomClassInData
    {
        public string ClassBField1 { get; set; }

        public string ClassBField2 { get; set; }

        public ClassC ClassC { get; set; }
    }

    public class ClassC
    {
        public string ClassCField1 { get; set; }

        public string ClassCField2 { get; set; }

        public string ClassCField3 { get; set; }
    }
}
