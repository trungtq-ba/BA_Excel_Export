using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace BAExcelExport
{

    /// <summary>
    /// Thông tin cột mà client phải truyền cho API để export Excel
    /// Nếu null hoặc không truyền xuống thì phía API sẽ lấy theo mặc định
    /// </summary>
    /// <Modified>
    /// Name     Date         Comments
    /// trungtq  11/12/2020   created
    /// </Modified>
    [Serializable]
    public class ColumnInfo
    {
        /// <summary>
        /// Tên của cột, trùng với tên của thuộc tính
        /// Tên của thuộc tính của đối tượng trả về từ API, dùng để mapping khi binding lại tiêu đề
        /// Nếu tên thuộc tính không trùng với tên thuộc tính phía Excel thì sẽ không mapping được
        /// </summary>
        public string ColumnName { get; set; }

        /// <summary>
        /// Nhãn hiển thị trên header
        /// Tiêu đề của cột khi export ra file Excel, Dịch theo culture của user trước khi truyền xuống
        /// Dựa vào PropertyName để mapping khi xuất excel, nếu không mappig được thì nó sẽ lấy theo tên thuộc tính phía Server (Tiếng Anh)
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
        public string Format { get; set; }

        /// <summary>
        /// Độ rộng cột
        /// Nếu không truyền thì mặc định là 100
        /// </summary>
        public int Width { get; set; }

        public ColumnInfo() { }

        public ColumnInfo(string columnName, string caption, bool visible, string format, int width)
        {
            this.ColumnName = columnName;
            this.Caption = caption;
            this.Visible = visible;
            this.Format = format;
            this.Width = width;
        }
    }
}
