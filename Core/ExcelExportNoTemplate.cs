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

                for (int i = 0; i < this.SettingColumns.Count; i++)
                {
                    ICell cell = row.CreateCell(i);
                    cell.SetCellValue(this.ReportTitle);
                }
            }
            if (!string.IsNullOrEmpty(this.ReportSubtitleLevel1))
            {
                IRow row = this.Sheet.CreateRow(0);

                for (int i = 0; i < this.SettingColumns.Count; i++)
                {
                    ICell cell = row.CreateCell(i);
                    cell.SetCellValue(this.ReportSubtitleLevel1);
                }
            }
            if (!string.IsNullOrEmpty(this.ReportSubtitleLevel2))
            {
                IRow row = this.Sheet.CreateRow(0);

                for (int i = 0; i < this.SettingColumns.Count; i++)
                {
                    ICell cell = row.CreateCell(i);
                    cell.SetCellValue(this.ReportSubtitleLevel2);
                }
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
                sheetRow = this.Sheet.CreateRow(i + 1);

                for (int j = 0; j < propertyInfos.Length; j++)
                {

                    ICell cellRow = sheetRow.CreateCell(j);

                    object cellvalue = propertyInfos[j].GetValue(this.DataSource[i], null);

                    if (cellvalue != null)
                    {

                        // Kiểm tra giá trị có là số không?
                        if (cellvalue.IsNumeric())
                        {
                            cellRow.SetCellValue(string.Format(new CultureInfo("en-US"), "{0:" + formatPart + "}", cellvalue));
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
