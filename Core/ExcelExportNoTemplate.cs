using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
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

        }
    }
}
