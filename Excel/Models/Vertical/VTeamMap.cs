using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aurora.IO.Excel.Mappings;

namespace Excel.Models.Vertical
{
    public class VTeamMap : ExcelMap<VTeam>
    {
        protected override ExcelMappingDirectionType MappingDirection => ExcelMappingDirectionType.Vertical;

        protected override int Header => 0;

        public static VTeamMap Create()
        {
            var map = new VTeamMap();

            var type = typeof(VTeam);
            map.Mapping.Add(1, type.GetProperty(nameof(VTeam.Name)));
            map.Mapping.Add(2, type.GetProperty(nameof(VTeam.FoundationYear)));
            map.Mapping.Add(3, type.GetProperty(nameof(VTeam.Titles)));

            return map;
        }
    }
}
