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
            return $"{fileName}_{DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss")}.xlsx";
        }

        public static HttpResponseMessage GenerateReport<TEntity>(ReportSourceTemplate<TEntity> template) where TEntity : ReportDataModelBase
        {
            HttpResponseMessage response = null;
            try
            {
                ExcelExportTable<ReportSourceTemplate<TEntity>, TEntity> report = new ExcelExportTable<ReportSourceTemplate<TEntity>, TEntity>(template);

                response = report.RenderReport();

            }
            catch (Exception ex)
            {

            }

            return response;
        }
    }
}
