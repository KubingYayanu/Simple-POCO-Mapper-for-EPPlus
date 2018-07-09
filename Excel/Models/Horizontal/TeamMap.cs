using Aurora.IO.Excel.Mappings;

namespace Excel.Models.Horizontal
{
    public class TeamMap : ExcelMap<Team>
    {
        protected override int Header => 0;

        public static TeamMap Create()
        {
            var map = new TeamMap();

            var type = typeof(Team);
            map.Mapping.Add(1, type.GetProperty(nameof(Team.Name)));
            map.Mapping.Add(2, type.GetProperty(nameof(Team.FoundationYear)));
            map.Mapping.Add(3, type.GetProperty(nameof(Team.Titles)));

            return map;
        }
    }
}