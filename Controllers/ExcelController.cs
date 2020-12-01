using BAExcelExport.ExcelExport;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;

// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace BAExcelExport.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
        private const int MAX_RECORD = 1000;

        // GET api/Excel/excelExport
        [HttpGet]
        [Route("excelExport")]
        public ActionResult ExcelExport()
        {
            var dataInput = DataHelper.GenerateData(MAX_RECORD);


            ReportSourceTemplate<ReportDataModel> template = new ReportSourceTemplate<ReportDataModel>()
            {
                FileName = typeof(ReportDataModel).Name,
                SheetName = "SheetName",
                ReportList = dataInput,
                SettingColumns = null,
            };


            Stopwatch sw = new Stopwatch();
            sw.Start();

            var resultContent = ReportHelper.GenerateReport(template);

            string fileName = ReportHelper.GetFileName(template.FileName);

            sw.Stop();

            return File(resultContent.Content.ReadAsByteArrayAsync().Result, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", sw.ElapsedMilliseconds + "_" + fileName);
        }

    }
}
