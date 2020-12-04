using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace BAExcelExport
{
    /// <summary>
    /// Class phục vụ việc Export dữ liệu ra file Excel
    /// Các định dạng export phải kế thừa và cài đặt lại 1 số method abstract
    /// </summary>
    /// <Modified>
    /// Name     Date         Comments
    /// trungtq  1/12/2020   created
    /// </Modified>
    public class ExcelExportBase<TSourceTemplate, TEntity>
        where TSourceTemplate : ReportSourceTemplate<TEntity>
        where TEntity : ReportDataModelBase
    {
        /// <summary>
        /// Tên Sheet của File Excel Export (mặc định khi không truyền vào) 
        /// </summary>
        /// <Modified>
        /// Name     Date         Comments
        /// trungtq  2/12/2020   created
        /// </Modified>
        private const string DefaultSheetName = "BAGPS";

        private const string DefaultFontName = "Times New Roman";

        private IWorkbook _Workbook = null;

        protected IWorkbook Workbook
        {
            get
            {
                if (_Workbook == null)
                {
                    _Workbook = new XSSFWorkbook();
                }
                return _Workbook;
            }
        }

        private ISheet _Sheet = null;

        protected ISheet Sheet
        {
            get
            {
                if (_Sheet == null)
                {
                    _Sheet = this.Workbook.CreateSheet(this.SheetName);
                }
                return _Sheet;
            }
        }

        /// <summary>
        /// Tên Sheet của File Excel Export ra 
        /// </summary>
        /// <Modified>
        /// Name     Date         Comments
        /// trungtq  2/12/2020   created
        /// </Modified>
        protected string SheetName { get; set; }

        /// <summary>
        /// Tên File Excel Export ra 
        /// </summary>
        /// <Modified>
        /// Name     Date         Comments
        /// trungtq  2/12/2020   created
        /// </Modified>
        protected string FileName { get; set; }

        /// <summary>
        /// Nguồn dữ liệu
        /// </summary>
        protected List<TEntity> DataSource { get; set; }

        protected List<ColumnInfo> SettingColumns { get; set; }

        protected TSourceTemplate SourceTemplate { get; set; }

        /// <summary>
        /// Tiêu đề của báo cáo
        /// </summary>
        protected string ReportTitle { get; set; }

        /// <summary>
        /// Noi dung: Nội dung: Thời gian vi phạm trên {Minutes} phút
        /// </summary>
        protected string ReportSubtitleLevel1 { get; set; }

        /// <summary>
        /// Ngày báo cáo
        /// </summary>
        protected string ReportSubtitleLevel2 { get; set; }

        private int _StartLoopRowIndex = -1;

        /// <summary>
        /// Dòng đầu tiên bind dữ liệu bảng
        /// </summary>
        protected int StartLoopRowIndex
        {
            get
            {
                if (_StartLoopRowIndex < 0)
                {
                    _StartLoopRowIndex = 0;

                    if (!string.IsNullOrEmpty(ReportTitle))
                    {
                        _StartLoopRowIndex += 1;
                    }
                    if (!string.IsNullOrEmpty(ReportSubtitleLevel1))
                    {
                        _StartLoopRowIndex += 1;
                    }
                    if (!string.IsNullOrEmpty(ReportSubtitleLevel2))
                    {
                        _StartLoopRowIndex += 1;
                    }
                }
                return _StartLoopRowIndex;
            }
        }

        /// <summary>
        /// Dòng cuối cùng bind dữ liệu bảng
        /// </summary>
        protected int EndLoopRowIndex
        {
            get
            {
                return StartLoopRowIndex + (DataSource != null ? DataSource.Count : 0);
            }
        }

        protected void ProcessSettingColumns(List<ColumnInfo> columnInfos)
        {
            if (columnInfos != null && columnInfos.Count > 0)
            {
                this.SettingColumns = columnInfos;
            }
            else
            {
                columnInfos = new List<ColumnInfo>();
                PropertyInfo[] propertyInfos = typeof(TEntity).GetProperties();

                foreach (PropertyInfo prop in propertyInfos)
                {
                    columnInfos.Add(new ColumnInfo()
                    {
                        ColumnName = prop.Name,
                        Caption = Regex.Replace(prop.Name, "([A-Z])", " $1").Trim(),
                        ColumnType = Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType

                    });
                }

                this.SettingColumns = columnInfos;
            }
        }

        public ExcelExportBase(TSourceTemplate template)
        {
            this.SourceTemplate = template;

            this.FileName = template.FileName;
            this.SheetName = template.SheetName ?? DefaultSheetName;
            this.DataSource = template.ReportList;
            this.ReportTitle = template.ReportTitle;
            this.ReportSubtitleLevel1 = template.ReportSubtitleLevel1;
            this.ReportSubtitleLevel2 = template.ReportSubtitleLevel2;
            this.ProcessSettingColumns(template.SettingColumns);
        }

        protected ICellStyle CreateCellStyle(HorizontalAlignment hAlign, VerticalAlignment vAlign, string fontName, double fontSize, bool isBold = false)
        {
            ICellStyle cellStyle = this.Workbook.CreateCellStyle();
            cellStyle.Alignment = hAlign;
            cellStyle.VerticalAlignment = vAlign;

            IFont font = CreateCellFont(fontName, fontSize, isBold);
            cellStyle.SetFont(font);

            return cellStyle;
        }

        protected ICellStyle CreateCellStyleReportTitle()
        {
            return CreateCellStyle(HorizontalAlignment.Center, VerticalAlignment.Center, DefaultFontName, 14f, true);
        }

        protected ICellStyle CreateCellStyleReportSubtitleLevel1()
        {
            return CreateCellStyle(HorizontalAlignment.Center, VerticalAlignment.Center, DefaultFontName, 12f, true);
        }

        protected ICellStyle CreateCellStyleReportSubtitleLevel2()
        {
            return CreateCellStyle(HorizontalAlignment.Center, VerticalAlignment.Center, DefaultFontName, 9f, true);
        }

        protected IFont CreateCellFont(string fontName, double fontSize, bool isBold = false)
        {
            var font = this.Workbook.CreateFont();
            font.IsBold = isBold;
            font.FontName = "Tahoma";
            font.FontHeightInPoints = fontSize;
            return font;
        }

        protected void AutoSizeColumn(bool autosize = false)
        {
            // It's heavy, it slows down your Excel if you have large data           
            if (autosize)
            {
                for (var i = 0; i < SettingColumns.Count; i++)
                {
                    this.Sheet.AutoSizeColumn(i);
                }
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

        /// <summary>
        /// Render ra Header của báo cáo
        /// </summary>
        /// <Modified>
        /// Name     Date         Comments
        /// trungtq  2/12/2020   created
        /// </Modified>
        protected virtual void RenderHeader() { }

        protected virtual void RenderBody()
        {

        }

        /// <summary>
        /// Render ra Summary của báo cáo
        /// </summary>
        /// <Modified>
        /// Name     Date         Comments
        /// trungtq  2/12/2020   created
        /// </Modified>
        protected virtual void RenderSummary()
        {

        }

        /// <summary>
        /// Render ra Footer của báo cáo
        /// </summary>
        /// <Modified>
        /// Name     Date         Comments
        /// trungtq  2/12/2020   created
        /// </Modified>
        protected virtual void RenderFooter()
        {

        }

        /// <summary>
        /// Render ra toàn bộ báo cáo.
        /// </summary>
        /// <Modified>
        /// Name     Date         Comments
        /// trungtq  2/12/2020   created
        /// </Modified>
        public HttpResponseMessage RenderReport()
        {
            HttpResponseMessage response = null;
            try
            {
                this.RenderHeader();

                this.RenderBody();

                this.RenderSummary();

                this.RenderFooter();

                this.AutoSizeColumn(true);

                // Tính lại độ rộng của cột
                this.CalculateColumnWidth();

                using (var memoryStream = new MemoryStream())
                {
                    this.Workbook.Write(memoryStream);

                    response = new HttpResponseMessage(HttpStatusCode.OK)
                    {
                        Content = new ByteArrayContent(memoryStream.ToArray())
                    };

                    response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                }
            }
            catch (Exception ex)
            {

            }
            return response;
        }
    }
}

