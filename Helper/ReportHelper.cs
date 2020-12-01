using BAExcelExport.ExcelExport;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;

namespace BAExcelExport
{
    public class ReportHelper
    {
        public static string GetFileName(string fileName)
        {
            return $"{fileName}_{DateTime.Now.ToString("yyyyMMddHHmmss")}.xlsx";
        }

        public static HttpResponseMessage GenerateReport<TEntity>(ReportSourceTemplate<TEntity> template) where TEntity : ReportDataModelBase
        {
            HttpResponseMessage response = null;
            try
            {
                DataExportTable<TEntity> report = new DataExportTable<TEntity>(template.ReportList, template.SettingColumns, template.FileName, template.SheetName);

                response = report.RenderReport();

            }
            catch (Exception ex)
            {

            }

            return response;
        }
    }
}
