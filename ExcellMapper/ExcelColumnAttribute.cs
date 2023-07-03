using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcellMapper
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelColumnAttribute : Attribute
    {
        public string ColumnName { get; set; }
        public byte ColumnIndex { get; set; }
        public object DefaultValueIfEmpty { get; set; } = null;

        public ExcelColumnAttribute(string columnName, object defaultValueIfEmpty = null)
        {
            ColumnName = columnName;
            this.DefaultValueIfEmpty = defaultValueIfEmpty;
        }
        public ExcelColumnAttribute(byte columnIndex, object defaultValueIfEmpty = null)
        {
            ColumnIndex = columnIndex;
            this.DefaultValueIfEmpty = defaultValueIfEmpty;
        }
    }
}
