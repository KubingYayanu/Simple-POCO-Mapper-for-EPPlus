using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aurora.IO.Excel.Mappings;
using Aurora.IO.Excel.Mappings.Attributes;

namespace Excel.Models.Vertical
{
    [ExcelMapper(MappingDirection = ExcelMappingDirectionType.Vertical)]
    public class VTeam
    {
        public string Name { get; set; }

        public int FoundationYear { get; set; }

        public int? Titles { get; set; }
    }
}
