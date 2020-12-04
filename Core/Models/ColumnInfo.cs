using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace BAExcelExport
{
    /// <summary>
    /// Thông tin cột cần lấy
    /// </summary>
    [Serializable]
    public class ColumnInfo
    {
        /// <summary>
        /// Thứ tự của cột khi export
        /// </summary>
        public int ColumnIndex { get; set; }

        /// <summary>
        /// Tên của cột, trùng với tên của thuộc tính
        /// </summary>
        public string ColumnName { get; set; }

        /// <summary>
        /// Kiểu dữ liệu của cột cần export
        /// </summary>
        public Type ColumnType { get; set; }

        /// <summary>
        /// Nhãn hiển thị trên header
        /// </summary>
        public string Caption { get; set; }

        /// <summary>
        /// Cột ẩn hay hiện
        /// </summary>
        public bool Visible { get; set; }

        public object GroupField { get; set; }

        public object HeaderMerge { get; set; }

        /// <summary>
        /// Định dạng cột
        /// </summary>
        public string Format { get; set; }

        /// <summary>
        /// Độ rộng cột
        /// </summary>
        public int Width { get; set; }

        /// <summary>
        /// Dữ liệu hiển thị trên Footer
        /// </summary>
        public object FooterData { get; set; }

        /// <summary>
        /// Thuộc tính mở rộng
        /// </summary>
        public IDictionary<string, object> ExtendedProperties { get; set; }

        /// <summary>
        /// Lấy dữ liệu từ từ điển mở rộng sao cho an toàn.
        /// </summary>
        /// <param name="propertyName"></param>
        /// <returns></returns>
        public object GetExtendedProperty(string propertyName)
        {
            if (!string.IsNullOrEmpty(propertyName))
            {
                if (this.ExtendedProperties != null && this.ExtendedProperties.ContainsKey(propertyName))
                {
                    return this.ExtendedProperties[propertyName];
                }
            }
            return null;
        }
    }
}
