namespace ExcelLite.Attributes
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = true)]
    public class GroupColumnNameAttribute : Attribute
    {
        public string _name;
        public int _depth;

        public GroupColumnNameAttribute(string name, int depth)
        {
            _name = name;
            _depth = depth;
        }
    }
}
