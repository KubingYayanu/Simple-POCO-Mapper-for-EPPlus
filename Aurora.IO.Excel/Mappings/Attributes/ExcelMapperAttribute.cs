using System;

namespace Aurora.IO.Excel.Mappings.Attributes
{
    [AttributeUsage(AttributeTargets.Class)]
    public class ExcelMapperAttribute : Attribute
    {
        public ExcelMappingDirectionType MappingDirection { get; set; } = ExcelMappingDirectionType.Horizontal;

        public int Header { get; set; } = 1;
    }
}