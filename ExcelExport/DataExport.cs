using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace BAExcelExport.ExcelExport
{
    public class DataExport : DataExportBase
    {
        public DataExport()
        {
            _headers = new List<string>();
            _type = new List<string>();
        }



        public override void WriteData<T>(List<T> exportData)
        {
            PropertyInfo[] propertyInfos = typeof(T).GetProperties();

            foreach (PropertyInfo prop in propertyInfos)
            {
                var type = Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType;
                _type.Add(type.Name);

                //space seperated name by caps for header
                string name = Regex.Replace(prop.Name, "([A-Z])", " $1").Trim();
                _headers.Add(name);
            }

            IRow sheetRow = null;

            ICellStyle CellCentertTopAlignment = _workbook.CreateCellStyle();
            CellCentertTopAlignment.Alignment = HorizontalAlignment.Center;
            CellCentertTopAlignment.VerticalAlignment = VerticalAlignment.Center;

            for (int i = 0; i < exportData.Count; i++)
            {
                sheetRow = _sheet.CreateRow(i + 1);

                for (int j = 0; j < propertyInfos.Length; j++)
                {

                    ICell cellRow = sheetRow.CreateCell(j);

                    object cellvalue = propertyInfos[j].GetValue(exportData[i], null);

                    if (cellvalue != null)
                    {
                        if (_type[j].ToLower() == "string")
                        {
                            cellRow.SetCellValue(cellvalue.ToString());
                        }
                        else if (_type[j].ToLower() == "int32")
                        {
                            cellRow.SetCellValue(Convert.ToInt32(cellvalue));
                        }
                        else if (_type[j].ToLower() == "double")
                        {
                            cellRow.SetCellValue(Convert.ToDouble(cellvalue));
                        }
                        else if (_type[j].ToLower() == "datetime")
                        {
                            cellRow.SetCellValue(Convert.ToDateTime(cellvalue).ToString("dd/MM/yyyy hh:mm:ss"));
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
