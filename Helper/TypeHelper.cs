using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace BAExcelExport
{
    public static class TypeHelper
    {
        /// <summary>
        /// Kiem tra doi tuong truyen vao co phai la so ko?
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <Modified>
        /// Name     Date         Comments
        /// trungtq  29/11/2020   created
        /// </Modified>
        public static bool IsNumeric(this object obj)
        {
            Type type = obj.GetType();

            switch (Type.GetTypeCode(type))
            {
                case TypeCode.Byte:
                case TypeCode.SByte:
                case TypeCode.UInt16:
                case TypeCode.UInt32:
                case TypeCode.UInt64:
                case TypeCode.Int16:
                case TypeCode.Int32:
                case TypeCode.Int64:
                case TypeCode.Decimal:
                case TypeCode.Double:
                case TypeCode.Single:
                    return true;
                case TypeCode.Object:
                    if (type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>))
                    {
                        return Nullable.GetUnderlyingType(type).IsNumeric();
                    }
                    return false;
                default:
                    return false;
            }
        }
    }
}
