using FluentExcel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
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

        private string DefaultFormatBool = "BOOLEAN";

        private string DefaultFormatInt = "#,##0";

        private string DefaultFormatDouble = "#,##0.###";

        private string DefaultFormatDatetime = "HH:mm:ss dd/MM/yyyy";

        protected PropertyInfo[] EntityProperties
        {
            get
            {
                return typeof(TEntity).GetProperties(BindingFlags.Public | BindingFlags.Instance | BindingFlags.GetProperty);
            }
        }

        private ExcelSetting ExcelSetting = Excel.Setting;

        private static IFormulaEvaluator _formulaEvaluator;

        private IWorkbook _Workbook = null;

        protected IWorkbook Workbook
        {
            get
            {
                if (_Workbook == null)
                {
                    var workbook = new XSSFWorkbook();

                    _formulaEvaluator = new XSSFFormulaEvaluator(workbook);
                    var props = workbook.GetProperties();
                    props.CoreProperties.Creator = ExcelSetting.Author;
                    props.CoreProperties.Subject = ExcelSetting.Subject;
                    props.ExtendedProperties.GetUnderlyingProperties().Company = ExcelSetting.Company;

                    _Workbook = workbook;
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

        private Dictionary<string, ColumnInfo> _SettingColumns = null;

        /// <summary>
        /// Từ điển cẩu hình
        /// </summary>
        /// <value>
        ///  Key: Tên thuộc tính
        ///  Value: Đối tượng cấu hình
        /// </value>
        /// <Modified>
        /// Name     Date         Comments
        /// trungtq  12/11/2020   created
        /// </Modified>
        protected Dictionary<string, ColumnInfo> SettingColumns
        {
            get
            {
                if (_SettingColumns == null)
                {
                    _SettingColumns = new Dictionary<string, ColumnInfo>();
                }
                return _SettingColumns;
            }
        }

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

        protected void ProcessSettingColumns(List<ColumnInfo> columnInfos)
        {
            if (columnInfos != null && columnInfos.Count > 0)
            {
                columnInfos.ForEach(item =>
                {
                    if (this.SettingColumns.ContainsKey(item.ColumnName))
                    {
                        this.SettingColumns.Add(item.ColumnName, item);
                    }
                    else
                    {
                        this.SettingColumns[item.ColumnName] = item;
                    }
                });
            }
            else
            {
                // TODO: can static properties or only instance properties?
                var propertyInfos = this.EntityProperties;

                foreach (PropertyInfo prop in propertyInfos)
                {
                    var columninfo = (new ColumnInfo()
                    {
                        ColumnName = prop.Name,
                        Caption = Regex.Replace(prop.Name, "([A-Z])", " $1").Trim(),
                        Visible = true
                    });

                    if (this.SettingColumns.ContainsKey(columninfo.ColumnName))
                    {
                        this.SettingColumns.Add(columninfo.ColumnName, columninfo);
                    }
                    else
                    {
                        this.SettingColumns[columninfo.ColumnName] = columninfo;
                    }
                }
            }
        }

        /// <summary>
        /// Đếm số cột được phép hiện
        /// </summary>
        /// <value>
        /// The visible column count.
        /// </value>
        /// <Modified>
        /// Name     Date         Comments
        /// trungtq  19/1/2021   created
        /// </Modified>
        protected int VisibleSettingColumnCount
        {
            get
            {
                return this.SettingColumns.Where(item => item.Value.Visible).Count();
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
            cellStyle.WrapText = true;

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
            return CreateCellStyle(HorizontalAlignment.Center, VerticalAlignment.Center, DefaultFontName, 9f, false);
        }

        protected ICellStyle CreateCellStyleTableHeader()
        {
            ICellStyle cell = CreateCellStyle(HorizontalAlignment.Center, VerticalAlignment.Center, DefaultFontName, 10f, true);

            cell.BorderBottom = BorderStyle.Thin;
            cell.BorderRight = BorderStyle.Thin;
            cell.BorderTop = BorderStyle.Thin;
            cell.BorderLeft = BorderStyle.Thin;

            return cell;
        }

        protected ICellStyle CreateCellStyleTableCell()
        {
            return CreateCellStyleTableCell(string.Empty);
        }

        protected Dictionary<string, ICellStyle> _DicColumnCellStyles = null;

        protected Dictionary<string, ICellStyle> DicColumnCellStyles
        {
            get
            {
                if (_DicColumnCellStyles == null)
                {
                    _DicColumnCellStyles = new Dictionary<string, ICellStyle>();

                    // TODO: can static properties or only instance properties?
                    var properties = this.EntityProperties;

                    if (properties != null && properties.Length > 0)
                    {
                        foreach (var item in properties)
                        {
                            // Kiểm tra giá trị có là số không?
                            if (item.PropertyType == typeof(bool))
                            {
                                _DicColumnCellStyles.Add(item.Name, this.CreateCellStyleTableCell(this.DefaultFormatBool));
                            }
                            else if (item.PropertyType == typeof(int))
                            {
                                _DicColumnCellStyles.Add(item.Name, this.CreateCellStyleTableCell(this.DefaultFormatInt));

                            }
                            else if (item.PropertyType == typeof(double))
                            {
                                _DicColumnCellStyles.Add(item.Name, this.CreateCellStyleTableCell(this.DefaultFormatDouble));
                            }
                            else if (item.PropertyType == typeof(DateTime))
                            {
                                _DicColumnCellStyles.Add(item.Name, this.CreateCellStyleTableCell(this.DefaultFormatDatetime));
                            }
                            else
                            {
                                _DicColumnCellStyles.Add(item.Name, this.CreateCellStyleTableCell());
                            }
                        }
                    }
                }

                return _DicColumnCellStyles;
            }
        }

        protected ICellStyle CreateCellStyleTableCell(string format)
        {
            ICellStyle cell = CreateCellStyle(HorizontalAlignment.Center, VerticalAlignment.Center, DefaultFontName, 9f, false);

            if (!string.IsNullOrEmpty(format))
            {
                cell.DataFormat = this.Workbook.CreateDataFormat().GetFormat(format);
            }

            cell.BorderBottom = BorderStyle.Thin;
            cell.BorderRight = BorderStyle.Thin;
            cell.BorderTop = BorderStyle.Thin;
            cell.BorderLeft = BorderStyle.Thin;

            return cell;
        }

        protected IFont CreateCellFont(string fontName, double fontSize, bool isBold = false)
        {
            var font = this.Workbook.CreateFont();
            font.IsBold = isBold;
            font.FontName = fontName;
            font.FontHeightInPoints = fontSize;
            return font;
        }

        private IDictionary<string, PropertyConfiguration> _PropertyConfigurations = null;

        /// <summary>
        /// Từ điển chưa thông tin cấu kình
        /// Key: Tên của thuộc tính/cột
        /// Value: giá trị cấu hình
        /// </summary>
        /// <Modified>
        /// Name     Date         Comments
        /// trungtq  19/1/2021   created
        /// </Modified>
        protected IDictionary<string, PropertyConfiguration> PropertyConfigurations
        {
            get
            {
                if (_PropertyConfigurations == null)
                {
                    _PropertyConfigurations = new Dictionary<string, PropertyConfiguration>();
                }
                return _PropertyConfigurations;
            }
        }

        /// <summary>
        /// Có sử dụng cấu hình không?, mặc định là có
        /// </summary>
        /// <Modified>
        /// Name     Date         Comments
        /// trungtq  19/1/2021   created
        /// </Modified>
        protected bool FluentConfigEnabled { get; set; } = false;

        protected IFluentConfiguration FluentConfig { get; set; } = null;

        protected virtual void PrepareFluentConfiguration()
        {
            // TODO: can static properties or only instance properties?
            var properties = this.EntityProperties;

            // Lấy thông tin từ fluent config
            if (this.ExcelSetting.FluentConfigs.TryGetValue(typeof(TEntity), out var fluentConfig))
            {
                this.FluentConfigEnabled = true;

                // adjust the auto index.
                (fluentConfig as FluentConfiguration<TEntity>)?.AdjustAutoIndex();

                this.FluentConfig = fluentConfig as FluentConfiguration<TEntity>;
            }

            // Duyệt qua các thuộc tính và truyền giá trị từ cấu hình xuống cho fluent config
            for (var i = 0; i < properties.Length; i++)
            {
                var property = properties[i];

                // Lấy thông tin từ fluent config
                if (this.FluentConfigEnabled && fluentConfig.PropertyConfigurations.TryGetValue(property.Name, out var pc))
                {
                    // Gán Header của cột và gán giá trị ẩn hiện cột.
                    if (this.SettingColumns.ContainsKey(property.Name))
                    {
                        pc.HasExcelTitle(this.SettingColumns[property.Name].Caption);
                        pc.IsExportIgnored = !this.SettingColumns[property.Name].Visible;

                        // Lấy cấu hình truyền vào từ ColumnInfo.
                        if (!string.IsNullOrEmpty(this.SettingColumns[property.Name].Format))
                        {
                            pc.HasDataFormatter(this.SettingColumns[property.Name].Format);
                        }
                    }

                    if (this.PropertyConfigurations.ContainsKey(property.Name))
                    {
                        this.PropertyConfigurations.Add(property.Name, pc);
                    }
                    else
                    {
                        this.PropertyConfigurations[property.Name] = pc;
                    }

                }
                else
                {
                    this.PropertyConfigurations.Add(property.Name, null);
                }
            }
        }

        protected virtual void ProcessMergeCell()
        {
            // merge cells
            var mergableConfigs = this.PropertyConfigurations.Values.Where(c => c != null && c.AllowMerge).ToList();
            if (mergableConfigs.Any())
            {

            }

        }

        protected string GetCellPosition(int row, int col)
        {
            col = Convert.ToInt32('A') + col;
            row = row + 1;
            return ((char)col) + row.ToString();
        }

        protected void AutoSizeColumn()
        {
            // It's heavy, it slows down your Excel if you have large data           
            if (this.ExcelSetting.AutoSizeColumnsEnabled)
            {
                for (var i = 0; i < this.SettingColumns.Count; i++)
                {
                    this.Sheet.AutoSizeColumn(i);
                }
            }
            else
            {
                // Duyệt qua và xử lý tất cả các cột cần ẩn.
                foreach (var pc in this.PropertyConfigurations)
                {
                    this.Sheet.SetColumnWidth(pc.Value.Index, this.SettingColumns[pc.Key].Width * 256);
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
            this.Sheet.PrintSetup.PaperSize = (short)PaperSize.A4;

            this.Sheet.PrintSetup.Landscape = this.SourceTemplate.Landscape;
        }

        /// <summary>
        /// Render ra Header của báo cáo
        /// </summary>
        /// <Modified>
        /// Name     Date         Comments
        /// trungtq  2/12/2020   created
        /// </Modified>
        protected virtual void RenderHeader()
        {
        }

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
        /// Xử lý ẩn hiện cột.
        /// </summary>
        /// <Modified>
        /// Name     Date         Comments
        /// trungtq  19/1/2021   created
        /// </Modified>
        protected virtual void ProcessVisibleColumn()
        {
            // Nếu có cột nào ẩn mới xử lý, không thì bỏ qua.
            if (this.PropertyConfigurations.Values.Any(v => v.IsExportIgnored == true))
            {
                // Duyệt qua và xử lý tất cả các cột cần ẩn.
                foreach (var pc in this.PropertyConfigurations)
                {
                    if (pc.Value.IsExportIgnored)
                    {
                        this.Sheet.SetColumnHidden(pc.Value.Index, true);
                    }
                }
            }
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
                this.PrepareFluentConfiguration();

                this.RenderHeader();

                this.RenderBody();

                this.RenderSummary();

                this.RenderFooter();

                this.AutoSizeColumn();

                this.ProcessVisibleColumn();

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

