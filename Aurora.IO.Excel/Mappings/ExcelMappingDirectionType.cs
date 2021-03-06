﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aurora.IO.Excel.Mappings
{
    public enum ExcelMappingDirectionType
    {
        /// <summary>
        /// Horizontal corresponds with each column matching a property on an object.
        /// Each row resprents an object.
        /// Use in combination with the Column property of the ExcelMap attribute.
        /// </summary>
        Horizontal = 0,

        /// <summary>
        /// Vertical corresponds with each row matching a property on an object.
        /// Each column represents an object.
        /// Use in combination with the Row property of the ExcelMap attribute.
        /// </summary>
        Vertical = 1
    }
}
