namespace ExcelLite
{
    public class Workbook
    {
        public Workbook(IEnumerable<Sheet> sheets)
        {
            Sheets = sheets;
        }

        public IEnumerable<Sheet> Sheets { get; set; }
    }
}
