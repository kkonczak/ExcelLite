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

        public SheetVisibility Visibility { get; set; } = SheetVisibility.Visible;

        public SheetView View { get; } = new SheetView();

        public bool UseBorders { get; set; }

        public IEnumerable<object> Data { get; set; }
    }
}
