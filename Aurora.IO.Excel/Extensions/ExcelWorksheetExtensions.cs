using Aurora.IO.Excel.Mappings;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Reflection;

namespace Aurora.IO.Excel.Extensions
{
    public static class ExcelWorksheetExtensions
    {
        public static TItem GetRecord<TItem>(this ExcelWorksheet sheet, int rowOrColumn, ExcelMap<TItem> map = null)
            where TItem : class
        {
            if (sheet == null)
            {
                return null;
            }

            if (map == null)
            {
                map = GetMap<TItem>(sheet);
            }

            if (rowOrColumn <= map.Header ||
                (map.MappingDirection == ExcelMappingDirectionType.Horizontal && rowOrColumn > sheet.Dimension.End.Row) ||
                (map.MappingDirection == ExcelMappingDirectionType.Vertical && rowOrColumn > sheet.Dimension.End.Column))
            {
                return null;
            }

            return GetItem(sheet, rowOrColumn, map);
        }

        private static TItem GetItem<TItem>(ExcelWorksheet sheet, int rowOrColumn, ExcelMap<TItem> map)
            where TItem : class
        {
            var item = Activator.CreateInstance<TItem>();
            foreach (var mapping in map.Mapping)
            {
                if ((map.MappingDirection == ExcelMappingDirectionType.Horizontal && mapping.Key > sheet.Dimension.End.Column) ||
                    (map.MappingDirection == ExcelMappingDirectionType.Vertical && mapping.Key > sheet.Dimension.End.Row))
                {
                    throw new ArgumentOutOfRangeException(nameof(rowOrColumn),
                        $"Key {mapping.Key} is outside of the sheet dimension using direction {map.MappingDirection}");
                }
                var value = (map.MappingDirection == ExcelMappingDirectionType.Horizontal)
                    ? sheet.GetValue(rowOrColumn, mapping.Key)
                    : sheet.GetValue(mapping.Key, rowOrColumn);
                if (value != null)
                {
                    // Test nullable
                    var type = mapping.Value.PropertyType.IsValueType
                        ? Nullable.GetUnderlyingType(mapping.Value.PropertyType) ?? mapping.Value.PropertyType
                        : mapping.Value.PropertyType;
                    var convertedValue = (type == typeof(string))
                        ? value.ToString().Trim()
                        : Convert.ChangeType(value, type);
                    mapping.Value.SetValue(item, convertedValue);
                }
                else
                {
                    // Explicitly set null values to prevent properties being initialized with their default values
                    mapping.Value.SetValue(item, null);
                }
            }
            return item;
        }

        private static ExcelMap<TItem> GetMap<TItem>(ExcelWorksheet sheet)
            where TItem : class
        {
            var method = typeof(ExcelMap<TItem>).GetMethod("CreateMap", BindingFlags.Static | BindingFlags.NonPublic);
            if (method == null)
            {
                throw new ArgumentNullException(nameof(method), $"Method CreateMap not found on type {typeof(ExcelMap<TItem>)}");
            }
            method = method.MakeGenericMethod(typeof(ExcelMap<TItem>));

            var map = method.Invoke(null, new object[] { sheet }) as ExcelMap<TItem>;
            if (map == null)
            {
                throw new ArgumentNullException(nameof(map), $"Map {typeof(ExcelMap<TItem>)} could not be created");
            }
            return map;
        }

        public static List<TItem> GetRecords<TItem>(this ExcelWorksheet sheet, ExcelMap<TItem> map = null)
            where TItem : class
        {
            if (sheet == null)
            {
                return new List<TItem>();
            }

            if (map == null)
            {
                map = GetMap<TItem>(sheet);
            }

            var items = new List<TItem>();
            var start = map.Header + 1;
            var endDimension = map.MappingDirection == ExcelMappingDirectionType.Horizontal
                ? sheet.Dimension.End.Row
                : sheet.Dimension.End.Column;
            for (var rowOrColumn = start; rowOrColumn <= endDimension; rowOrColumn++)
            {
                var item = GetItem(sheet, rowOrColumn, map);
                items.Add(item);
            }

            return items;
        }
    }
}