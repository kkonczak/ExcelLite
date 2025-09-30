namespace ExcelLite.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ColumnPositionAttribute : Attribute
    {
        public int _index;

        public ColumnPositionAttribute(int index)
        {
            _index = index;
        }
    }
}
