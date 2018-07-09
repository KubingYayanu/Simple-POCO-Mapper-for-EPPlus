using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aurora.IO.Excel.Mappings;
using Aurora.IO.Excel.Mappings.Attributes;

namespace Excel.Models.Vertical
{
    [ExcelMapper(MappingDirection = ExcelMappingDirectionType.Vertical, Header = 0)]
    public class VTeamAttributes
    {
        [ExcelMap(Row = 1)]
        public string Name { get; set; }

        [ExcelMap(Row = 2)]
        public int FoundationYear { get; set; }

        [ExcelMap(Row = 3)]
        public int? Titles { get; set; }
    }
}
