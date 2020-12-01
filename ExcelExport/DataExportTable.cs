using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;

namespace BAExcelExport.ExcelExport
{
    public class DataExportTable<TEntity> : DataExportBase<TEntity> where TEntity : ReportDataModelBase
    {
        public DataExportTable(List<TEntity> data, List<ColumnInfo> settingColumns, string fileName, string sheetName) : base(data, settingColumns, fileName, sheetName)
        {
        }

        protected override void RenderHeader(ISheet sheet, ICellStyle headerStyle)
        {
            //Header
            var header = sheet.CreateRow(0);
            for (var i = 0; i < SettingColumns.Count; i++)
            {
                var cell = header.CreateCell(i);
                cell.SetCellValue(SettingColumns[i].Caption);
                cell.CellStyle = headerStyle;
            }
        }

        protected override void RenderBody(IWorkbook workbook, ISheet sheet)
        {
            PropertyInfo[] propertyInfos = typeof(TEntity).GetProperties();

            IRow sheetRow = null;

            ICellStyle cellStyle = this.CreateCellStyle(workbook, HorizontalAlignment.Center, VerticalAlignment.Center);

            for (int i = 0; i < dataSource.Count; i++)
            {
                sheetRow = sheet.CreateRow(i + 1);

                for (int j = 0; j < this.SettingColumns.Count; j++)
                {

                    ICell cellRow = sheetRow.CreateCell(j);

                    object cellvalue = propertyInfos[j].GetValue(dataSource[i], null);

                    if (cellvalue != null)
                    {
                        if (cellvalue != null)
                        {
                            if (this.SettingColumns[j].ColumnType.Name.ToLower() == "string")
                            {
                                cellRow.SetCellValue(cellvalue.ToString());
                            }
                            else if (this.SettingColumns[j].ColumnType.Name.ToLower() == "int32")
                            {
                                cellRow.SetCellValue(Convert.ToInt32(cellvalue));
                            }
                            else if (this.SettingColumns[j].ColumnType.Name.ToLower() == "double")
                            {
                                cellRow.SetCellValue(Convert.ToDouble(cellvalue));
                            }
                            else if (this.SettingColumns[j].ColumnType.Name.ToLower() == "datetime")
                            {
                                cellRow.SetCellValue(Convert.ToDateTime(cellvalue).ToString("dd/MM/yyyy hh:mm:ss"));
                            }
                        }
                        else
                        {
                            cellRow.SetCellValue(string.Empty);
                        }
                    }
                    else
                    {
                        cellRow.SetCellValue(string.Empty);
                    }

                    cellRow.CellStyle = cellStyle;
                    cellRow.CellStyle.WrapText = true;
                }
            }
        }

    }
}
