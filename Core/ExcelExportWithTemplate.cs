using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace BAExcelExport
{
    public class ExcelExportWithTemplate
    {
        /// <summary>
        /// Đường dẫn file Template
        /// </summary>
        /// <Modified>
        /// Name     Date         Comments
        /// trungtq  1/12/2020   created
        /// </Modified>
        protected string FileTemplatePath { get; set; } = string.Empty;

        /// <summary>
        /// Có dùng file Template không?
        /// Nếu không truyền vào đường dẫn FileTemplate => sinh tự động, không cần file Template.
        /// </summary>
        /// <Modified>
        /// Name     Date         Comments
        /// trungtq  2/12/2020   created
        /// </Modified>
        protected bool EnableFileTemplate
        {
            get
            {
                return !string.IsNullOrEmpty(FileTemplatePath);
            }
        }
    }
}
