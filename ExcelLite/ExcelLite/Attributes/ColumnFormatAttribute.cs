namespace ExcelLite.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ColumnFormatAttribute : Attribute
    {
        public string _format;

        public ColumnFormatAttribute(string format)
        {
            _format = format;
        }
    }
}
