using BAExcelExport.ExcelExport;
using BAExcelExport.Models;
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

            var exportExcel = new DataExport();

            Stopwatch sw = new Stopwatch();
            sw.Start();

            var resultContent = exportExcel.Export(dataInput, typeof(ReportDataModel).Name, "ReportSummary");

            sw.Stop();

            return File(resultContent.Content.ReadAsByteArrayAsync().Result, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", sw.ElapsedMilliseconds + "_" + exportExcel.FileName);
        }
    }
}
