using FluentExcel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;

namespace BAExcelExport
{
    /// <summary>
    /// Export ra Excel Không cần Template
    /// Chỉ dùng cho dạng export ra đơn giản
    /// </summary>
    /// <Modified>
    /// Name     Date         Comments
    /// trungtq  9/12/2020   created
    /// </Modified>
    public class ExcelExportNoTemplate<TSourceTemplate, TEntity> : ExcelExportBase<TSourceTemplate, TEntity>
            where TSourceTemplate : ReportSourceTemplate<TEntity>
            where TEntity : ReportDataModelBase
    {

        public ExcelExportNoTemplate(TSourceTemplate template) : base(template)
        {
        }

        protected override void RenderHeader()
        {
            if (!string.IsNullOrEmpty(this.ReportTitle))
            {
                IRow row = this.Sheet.CreateRow(this.Sheet.LastRowNum);
                ICell cell = row.CreateCell(0);
                cell.CellStyle = this.CreateCellStyleReportTitle();
                cell.SetCellValue(this.ReportTitle);
                var cra = new NPOI.SS.Util.CellRangeAddress(0, 0, 0, this.SettingColumns.Count);
                this.Sheet.AddMergedRegion(cra);
            }
            if (!string.IsNullOrEmpty(this.ReportSubtitleLevel1))
            {
                IRow row = this.Sheet.CreateRow(this.Sheet.LastRowNum + 1);
                ICell cell = row.CreateCell(0);
                cell.CellStyle = this.CreateCellStyleReportSubtitleLevel1();
                var cra = new NPOI.SS.Util.CellRangeAddress(row.RowNum, row.RowNum, 0, this.SettingColumns.Count);
                cell.SetCellValue(this.ReportSubtitleLevel1);
                this.Sheet.AddMergedRegion(cra);
            }
            if (!string.IsNullOrEmpty(this.ReportSubtitleLevel2))
            {
                IRow row = this.Sheet.CreateRow(this.Sheet.LastRowNum + 1);
                ICell cell = row.CreateCell(0);
                cell.CellStyle = this.CreateCellStyleReportSubtitleLevel2();
                var cra = new NPOI.SS.Util.CellRangeAddress(row.RowNum, row.RowNum, 0, this.SettingColumns.Count);
                cell.SetCellValue(this.ReportSubtitleLevel2);
                this.Sheet.AddMergedRegion(cra);
            }
        }



        protected override void RenderBody()
        {
            try
            {
                // TODO: can static properties or only instance properties?
                PropertyInfo[] propertyInfos = typeof(TEntity).GetProperties(BindingFlags.Public | BindingFlags.Instance | BindingFlags.GetProperty);

                // Render Table Header
                var headerRow = this.Sheet.CreateRow(this.Sheet.LastRowNum + 1);

                ICellStyle headerCellStyle = this.CreateCellStyleTableHeader();

                for (var i = 0; i < propertyInfos.Length; i++)
                {
                    var property = propertyInfos[i];

                    var config = this.PropertyConfigurations[property.Name];

                    // Mặc định: tiêu đề cột là tên thuộc tính của đối tượng.
                    var title = propertyInfos[i].Name;

                    int index = i;
                    if (config != null)
                    {
                        // Lấy giá trị title
                        if (!string.IsNullOrEmpty(config.Title))
                        {
                            title = config.Title;
                        }

                        // Nếu không cần export cột này thì next đến cột khác
                        if (config.IsExportIgnored) continue;

                        index = config.Index;

                        if (index < 0)
                            throw new Exception($"The excel cell index value cannot be less then '0' for the property: {property.Name}, see HasExcelIndex(int index) methods for more informations.");
                    }

                    var cell = headerRow.CreateCell(index);
                    cell.CellStyle = headerCellStyle;
                    cell.SetCellValue(title);
                }

                // Duyệt và binding dữ liệu
                for (int i = 0; i < this.DataSource.Count; i++)
                {
                    IRow sheetRow = this.Sheet.CreateRow(this.Sheet.LastRowNum + 1);

                    for (int j = 0; j < propertyInfos.Length; j++)
                    {
                        var property = propertyInfos[j];

                        var config = this.PropertyConfigurations[property.Name];

                        int index = j;

                        if (config != null)
                        {
                            // Nếu không cần export cột này thì next đến cột khác
                            if (config.IsExportIgnored) continue;

                            index = config.Index;

                            if (index < 0)
                                throw new Exception($"The excel cell index value cannot be less then '0' for the property: {property.Name}, see HasExcelIndex(int index) methods for more informations.");
                        }

                        ICell cell = sheetRow.CreateCell(index);

                        cell.CellStyle = this.ColumnCellStyles[j];

                        Type cellType = property.PropertyType.UnwrapNullableType();

                        object cellvalue = property.GetValue(this.DataSource[i], null);

                        if (cellvalue != null)
                        {
                            if (!string.IsNullOrEmpty(config?.Formatter) && cellvalue is IFormattable fv)
                            {
                                // the formatter isn't excel supported formatter, but it's a C# formatter.
                                // The result is the Excel cell data type become String.
                                cell.SetCellValue(fv.ToString(config.Formatter, CultureInfo.CurrentCulture));

                                continue;
                            }

                            // Kiểm tra giá trị có là số không?
                            if (cellType == typeof(bool))
                            {
                                cell.SetCellValue(Convert.ToBoolean(cellvalue));
                            }
                            else if (cellType.IsInteger())
                            {
                                cell.SetCellValue(Convert.ToInt32(cellvalue));
                            }
                            else if (cellType.IsDouble())
                            {
                                cell.SetCellValue(Convert.ToDouble(cellvalue));
                            }
                            else if (cellType == typeof(DateTime))
                            {
                                cell.SetCellValue(Convert.ToDateTime(cellvalue));
                            }
                            else
                            {
                                cell.SetCellValue(cellvalue.ToString());
                            }
                        }
                        else
                        {
                            cell.SetCellValue(string.Empty);
                        }
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }
    }
}
