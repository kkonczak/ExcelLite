namespace ExcelLite.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ColumnNameAttribute : Attribute
    {
        public string _name;

        public ColumnNameAttribute(string name)
        {
            _name = name;
        }
    }
}
