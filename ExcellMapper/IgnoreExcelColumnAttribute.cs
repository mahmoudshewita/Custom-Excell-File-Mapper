using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcellMapper
{
    [AttributeUsage(AttributeTargets.Property)]
    public class IgnoreExcelColumnAttribute : Attribute
    {
        public bool Ignore { get; set; }

        public IgnoreExcelColumnAttribute(bool ignoreColumn = true)
        {
            Ignore = ignoreColumn;
        }
    }
}
