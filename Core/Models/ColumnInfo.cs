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
        /// Mặc định: hiện
        /// </summary>
        public bool Visible { get; set; } = true;

        /// <summary>
        /// Định dạng cột
        /// </summary>
        public string DataFormat { get; set; }

        /// <summary>
        /// Độ rộng cột
        /// </summary>
        public int Width { get; set; } = 100;

        /// <summary>
        /// Công thức của cột
        /// </summary>
        /// <Modified>
        /// Name     Date         Comments
        /// trungtq  9/12/2020   created
        /// </Modified>
        public string Formula { get; set; }
       
    }
}
