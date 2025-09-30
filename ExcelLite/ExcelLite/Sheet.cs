namespace ExcelLite
{
    public class Sheet
    {
        public Sheet(string name, IEnumerable<object> data)
        {
            Name = name;
            Data = data;
        }

        public string Name { get; set; }

        public SheetView View { get; } = new SheetView();

        public IEnumerable<object> Data { get; set; }
    }
}
