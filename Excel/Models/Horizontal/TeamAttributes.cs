using Aurora.IO.Excel.Mappings.Attributes;

namespace Excel.Models.Horizontal
{
    [ExcelMapper(Header = 0)]
    public class TeamAttributes
    {
        [ExcelMap(Column = 1)]
        public string Name { get; set; }
        
        [ExcelMap(Column = 2)]
        public int FoundationYear { get; set; }
        
        [ExcelMap(Column = 3)]
        public int? Titles { get; set; }
    }
}