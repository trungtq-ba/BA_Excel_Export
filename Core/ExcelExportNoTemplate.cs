﻿using NPOI.SS.UserModel;
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
                IRow row = this.Sheet.CreateRow(0);
                ICell cell = row.CreateCell(0);
                cell.CellStyle = this.CreateCellStyleReportTitle();
                cell.SetCellValue(this.ReportTitle);
                var cra = new NPOI.SS.Util.CellRangeAddress(0, 0, 0, SettingColumns.Count);
                this.Sheet.AddMergedRegion(cra);
            }
            if (!string.IsNullOrEmpty(this.ReportSubtitleLevel1))
            {
                IRow row = this.Sheet.CreateRow(1);
                ICell cell = row.CreateCell(0);
                cell.CellStyle =this.CreateCellStyleReportSubtitleLevel1();
                var cra = new NPOI.SS.Util.CellRangeAddress(1, 1, 0, SettingColumns.Count);
                cell.SetCellValue(this.ReportSubtitleLevel1);
                this.Sheet.AddMergedRegion(cra);
            }
            if (!string.IsNullOrEmpty(this.ReportSubtitleLevel2))
            {
                IRow row = this.Sheet.CreateRow(2);
                ICell cell = row.CreateCell(0);
                cell.CellStyle = this.CreateCellStyleReportSubtitleLevel2();
                var cra = new NPOI.SS.Util.CellRangeAddress(2, 2, 0, SettingColumns.Count);
                cell.SetCellValue(this.ReportSubtitleLevel2);
                this.Sheet.AddMergedRegion(cra);
            }
        }

        protected override void RenderBody()
        {
            PropertyInfo[] propertyInfos = typeof(TEntity).GetProperties();

            // Render Table Header
            var headerRow = this.Sheet.CreateRow(3);

            ICellStyle headerCellStyle = this.CreateCellStyleTableHeader();

            for (var i = 0; i <this.SettingColumns.Count; i++)
            {
                var cell = headerRow.CreateCell(i);
                
                cell.CellStyle = headerCellStyle;


                cell.SetCellValue(this.SettingColumns[i].Caption);
            }

            // Render Table Body
            IRow sheetRow = null;

            ICellStyle CellCentertTopAlignment = this.CreateCellStyleTableCell();

            string formatPart = string.Empty;

            for (int i = 0; i < this.DataSource.Count; i++)
            {
                sheetRow = this.Sheet.CreateRow((StartLoopRowIndex +1)+i);

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
                }
            }
        }
    }
}
