namespace ExcelLite.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ColumnIgnoreAttribute : Attribute
    {
        public ColumnIgnoreAttribute()
        {
        }
    }
}
