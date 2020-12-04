using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;

namespace BAExcelExport
{
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
                IRow row = this.Sheet.CreateRow(0);
                ICell cell = row.CreateCell(0);
                cell.CellStyle.WrapText = true;
                cell.SetCellValue(this.ReportTitle);
                var cra = new NPOI.SS.Util.CellRangeAddress(0, 0, 0, SettingColumns.Count);
                this.Sheet.AddMergedRegion(cra);
            }
            if (!string.IsNullOrEmpty(this.ReportSubtitleLevel1))
            {
                IRow row = this.Sheet.CreateRow(1);
                ICell cell = row.CreateCell(0);
                cell.CellStyle.WrapText = true;
                var cra = new NPOI.SS.Util.CellRangeAddress(1, 1, 0, SettingColumns.Count);
                cell.SetCellValue(this.ReportSubtitleLevel1);
                this.Sheet.AddMergedRegion(cra);
            }
            if (!string.IsNullOrEmpty(this.ReportSubtitleLevel2))
            {
                IRow row = this.Sheet.CreateRow(2);
                ICell cell = row.CreateCell(0);
                cell.CellStyle.WrapText = true;
                var cra = new NPOI.SS.Util.CellRangeAddress(2, 2, 0, SettingColumns.Count);
                cell.SetCellValue(this.ReportSubtitleLevel2);
                this.Sheet.AddMergedRegion(cra);
            }
        }

        protected override void RenderBody()
        {
            PropertyInfo[] propertyInfos = typeof(TEntity).GetProperties();

            IRow sheetRow = null;

            ICellStyle CellCentertTopAlignment = this.Workbook.CreateCellStyle();
            CellCentertTopAlignment.Alignment = HorizontalAlignment.Center;
            CellCentertTopAlignment.VerticalAlignment = VerticalAlignment.Center;

            string formatPart = string.Empty;

            for (int i = 0; i < this.DataSource.Count; i++)
            {
                sheetRow = this.Sheet.CreateRow(StartLoopRowIndex+i);

                for (int j = 0; j < propertyInfos.Length; j++)
                {

                    ICell cellRow = sheetRow.CreateCell(j);

                    object cellvalue = propertyInfos[j].GetValue(this.DataSource[i], null);

                    if (cellvalue != null)
                    {
                        // Kiểm tra giá trị có là số không?
                        if (cellvalue.IsNumeric())
                        {
                            //cellRow.SetCellValue(string.Format(new CultureInfo("en-US"), "{0:" + formatPart + "}", cellvalue));
                            cellRow.SetCellType(CellType.Numeric);
                            cellRow.SetCellValue(cellvalue.ToString());
                        }
                        else
                        {
                            cellRow.SetCellValue(string.Format("{0:" + formatPart + "}", cellvalue));
                        }
                    }
                    else
                    {
                        cellRow.SetCellValue(string.Empty);
                    }

                    cellRow.CellStyle = CellCentertTopAlignment;
                    cellRow.CellStyle.WrapText = true;
                }
            }
        }
    }
}
