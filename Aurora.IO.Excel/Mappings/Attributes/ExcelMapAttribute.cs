using System;

namespace Aurora.IO.Excel.Mappings.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelMapAttribute : Attribute
    {
        /// <summary>
        /// Use column with HeaderDirection.Horizontal
        /// </summary>
        public int Column { get; set; }

        /// <summary>
        /// Use row with HeaderDirection.Vertical
        /// </summary>
        public int Row { get; set; }
    }
}