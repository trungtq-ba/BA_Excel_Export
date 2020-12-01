using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace BAExcelExport.ExcelExport
{
    public abstract class DataExportBase<TEntity> where TEntity:class
    {
        protected string _sheetName;
        protected string _fileName;
        private const string DefaultSheetName = "BAGPS";

        /// <summary>
        /// Danh sách cột cấu hình
        /// </summary>
        private List<ColumnInfo> _SettingColumns = null;

        protected List<ColumnInfo> SettingColumns
        {
            get
            {
                if (_SettingColumns == null || _SettingColumns.Count == 0)
                {
                    _SettingColumns = new List<ColumnInfo>();

                    PropertyInfo[] propertyInfos = typeof(TEntity).GetProperties();

                    foreach (PropertyInfo prop in propertyInfos)
                    {
                        _SettingColumns.Add(new ColumnInfo()
                        {
                            ColumnName = prop.Name,
                            Caption = Regex.Replace(prop.Name, "([A-Z])", " $1").Trim(),
                            ColumnType = Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType

                        });
                    }
                }

                return _SettingColumns;
            }
        }

        public string FileName
        {
            get
            {
                return $"{_fileName}_{DateTime.Now.ToString("yyyyMMddHHmmss")}.xlsx";
            }
        }

        public DataExportBase(List<TEntity> data, List<ColumnInfo> settingColumns, string fileName, string sheetName = DefaultSheetName)
        {
            _fileName = fileName;
            _sheetName = sheetName;
            this.dataSource = data;
            this._SettingColumns = settingColumns;
        }

        /// <summary>
        /// Nguồn dữ liệu
        /// </summary>
        protected List<TEntity> dataSource { get; set; }

        protected ICellStyle CreateCellStyle(IWorkbook workbook, HorizontalAlignment hAlign, VerticalAlignment vAlign)
        {
            return CreateCellStyle(workbook, hAlign, vAlign, false);
        }

        protected ICellStyle CreateCellStyle(IWorkbook workbook, HorizontalAlignment hAlign, VerticalAlignment vAlign, bool isBold)
        {
            ICellStyle cellStyle = workbook.CreateCellStyle();
            cellStyle.Alignment = hAlign;
            cellStyle.VerticalAlignment = vAlign;

            if (isBold)
            {
                var headerFont = workbook.CreateFont();
                headerFont.IsBold = true;
                cellStyle.SetFont(headerFont);
            }

            return cellStyle;
        }

        protected void AutoSizeColumn(ISheet sheet, bool autosize = false)
        {
            // It's heavy, it slows down your Excel if you have large data           
            if (autosize)
            {
                for (var i = 0; i < SettingColumns.Count; i++)
                {

                    sheet.AutoSizeColumn(i);
                }
            }
        }

        protected abstract void RenderHeader(ISheet sheet, ICellStyle headerStyle);

        
        protected abstract void RenderBody(IWorkbook workbook, ISheet sheet);
        

        protected virtual void RenderSummary()
        {

        }

        protected virtual void RenderFooter()
        {

        }

        public HttpResponseMessage RenderReport()
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet(_sheetName);

            ICellStyle headerStyle = CreateCellStyle(workbook, HorizontalAlignment.Center, VerticalAlignment.Center, true);

            this.RenderHeader(sheet, headerStyle);

            this.RenderBody(workbook, sheet);

            this.RenderSummary();
            this.RenderFooter();

            this.AutoSizeColumn(sheet, true);

            using (var memoryStream = new MemoryStream())
            {
                workbook.Write(memoryStream);
                var response = new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ByteArrayContent(memoryStream.ToArray())
                };

                response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

                return response;
            }
        }

        /// <summary>
        /// Tính toán lại độ rộng của cột.
        /// </summary>
        /// <Modified>
        /// Name     Date         Comments
        /// trungtq  27/02/2015   created
        /// </Modified>
        protected virtual void CalculateColumnWidth()
        {
        }

        public HttpResponseMessage Export()
        {
            // Tính lại độ rộng của cột
            this.CalculateColumnWidth();

            // Chạy báo cáo
            return this.RenderReport();
        }
    }
}
