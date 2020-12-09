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
                PropertyInfo[] propertyInfos = typeof(TEntity).GetProperties();

                // Render Table Header
                var headerRow = this.Sheet.CreateRow(this.Sheet.LastRowNum + 1);

                ICellStyle headerCellStyle = this.CreateCellStyleTableHeader();

                for (var i = 0; i < this.SettingColumns.Count; i++)
                {
                    var cell = headerRow.CreateCell(i);
                    cell.CellStyle = headerCellStyle;
                    cell.SetCellValue(this.SettingColumns[i].Caption);
                }

                for (int i = 0; i < this.DataSource.Count; i++)
                {
                    IRow sheetRow = this.Sheet.CreateRow(this.Sheet.LastRowNum + 1);

                    for (int j = 0; j < propertyInfos.Length; j++)
                    {
                        ICell cell = sheetRow.CreateCell(j);

                        Type cellType = propertyInfos[j].PropertyType;
                        object cellvalue = propertyInfos[j].GetValue(this.DataSource[i], null);

                        if (cellvalue != null)
                        {
                            // Kiểm tra giá trị có là số không?
                            if (cellType == typeof(bool))
                            {
                                cell.SetCellValue(Convert.ToBoolean(cellvalue));

                            }
                            else if (cellType == typeof(int))
                            {
                                cell.SetCellValue(Convert.ToInt32(cellvalue));
                            }
                            else if (cellType == typeof(double))
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

                        cell.CellStyle = this.ColumnCellStyles[j];
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }
    }
}
