using Aurora.IO.Excel.Mappings.Attributes;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace Aurora.IO.Excel.Mappings
{
    public class ExcelMap<TItem> where TItem : class
    {
        protected internal virtual ExcelMappingDirectionType MappingDirection { get; set; } =
            ExcelMappingDirectionType.Horizontal;

        protected internal virtual int Header { get; set; } = 1;

        protected internal Dictionary<int, PropertyInfo> Mapping { get; private set; } = new Dictionary<int, PropertyInfo>();

        protected static TMap CreateMap<TMap>(ExcelWorksheet sheet) where TMap : ExcelMap<TItem>
        {
            var map = Activator.CreateInstance<TMap>();

            var type = typeof(TItem);

            // Check if we map by attributes or by column header name
            var mapper = type.GetCustomAttribute<ExcelMapperAttribute>();
            if (mapper != null)
            {
                // Map by attribute
                map.MappingDirection = mapper.MappingDirection;
                map.Header = mapper.Header;

                type.GetProperties()
                    .Select(x => new { Property = x, Attribute = x.GetCustomAttribute<ExcelMapAttribute>() })
                    .Where(x => x.Attribute != null)
                    .ToList()
                    .ForEach(prop =>
                    {
                        var key = map.MappingDirection == ExcelMappingDirectionType.Horizontal
                            ? prop.Attribute.Column
                            : prop.Attribute.Row;
                        map.Mapping.Add(key, prop.Property);
                    });
            }

            if (!map.Mapping.Any())
            {
                // Map by column / row header name
                var props = type.GetProperties().ToList();

                // Determine end dimension for the header
                var endDimension = map.MappingDirection == ExcelMappingDirectionType.Horizontal
                    ? sheet.Dimension.End.Column
                    : sheet.Dimension.End.Row;
                for (var rowOrColumn = 1; rowOrColumn <= endDimension; rowOrColumn++)
                {
                    var parameter = map.MappingDirection == ExcelMappingDirectionType.Horizontal
                        ? sheet.GetValue<string>(map.Header, rowOrColumn)
                        : sheet.GetValue<string>(rowOrColumn, map.Header);
                    if (string.IsNullOrWhiteSpace(parameter))
                    {
                        var message = map.MappingDirection == ExcelMappingDirectionType.Horizontal
                            ? $"Column {rowOrColumn} has no parameter name"
                            : $"Row {rowOrColumn} has no parameter name";
                        throw new ArgumentNullException(nameof(parameter), message);
                    }

                    // Remove spaces
                    parameter = parameter.Replace(" ", string.Empty).Trim();

                    // Map to property
                    var prop = props.FirstOrDefault(x => StringComparer.OrdinalIgnoreCase.Equals(x.Name, parameter));
                    if (prop == null)
                    {
                        throw new ArgumentNullException(nameof(parameter), $"No property {parameter} found on type {typeof(TItem)}");
                    }
                    map.Mapping.Add(rowOrColumn, prop);
                }
            }

            return map;
        }
    }
}