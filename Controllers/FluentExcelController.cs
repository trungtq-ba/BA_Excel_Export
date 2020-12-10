using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;

namespace BAExcelExport.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class FluentExcelController : ControllerBase
    {
        private const int MAX_RECORD = 100;

        // GET api/Excel/excelExport
        [HttpGet]
        [Route("fluentExcelExport")]
        public ActionResult ExcelExport()
        {
            var dataInput = DataHelper.GenerateData(MAX_RECORD);


            ReportSourceTemplate<ReportDataModel> template = new ReportSourceTemplate<ReportDataModel>()
            {
                ReportTitle = "BÁO CÁO DANH SÁCH ĐIỂM",
                ReportSubtitleLevel1 = "TIÊU ĐỀ CON CỦA BÁO CÁO DANH SÁCH ĐIỂM",
                ReportSubtitleLevel2 = "MÔ TẢ CỦA BÁO CÁO DANH SÁCH ĐIỂM",
                FileName = typeof(ReportDataModel).Name,
                SettingColumns = null,
                ReportList = dataInput
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
