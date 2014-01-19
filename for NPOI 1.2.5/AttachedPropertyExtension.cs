using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.Runtime.CompilerServices;

namespace NPOI.CSS
{
    internal static class AttachedPropertyExtension
    {
        private static ConditionalWeakTable<IWorkbook, Dictionary<string, ICellStyle>> _table =
                            new ConditionalWeakTable<IWorkbook, Dictionary<string, ICellStyle>>();

        public static ICellStyle GetCellStyle(this IWorkbook owner, string propertyName)
        {
            Dictionary<string, ICellStyle> values;
            if (_table.TryGetValue(owner, out values))
            {
                ICellStyle temp;
                if (values.TryGetValue(propertyName, out temp))
                {
                    return temp;
                }
            }

            return null;
        }

        public static void AttachedCellStyle(this IWorkbook owner, string propertyName, ICellStyle value)
        {
            Dictionary<string, ICellStyle> values;
            if (!_table.TryGetValue(owner, out values))
            {
                values = new Dictionary<string, ICellStyle>();
                _table.Add(owner, values);
            }

            values[propertyName] = value;
        }
    }
}