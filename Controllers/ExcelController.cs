using FluentExcel;
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
            FluentConfiguration();

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

        /// <summary>
        /// Use fluent configuration api. (doesn't poison your POCO)
        /// </summary>
        private static void FluentConfiguration()
        {
            var fc = Excel.Setting.For<ReportDataModel>();
            //Excel.Setting.AutoSizeColumnsEnabled = false;

            fc.HasStatistics("Tổng", "SUM", 4, 5,6);

            fc.Property(r => r.OrderNumber)
              .HasExcelIndex(0)
              .HasExcelTitle("STT")
              .IsMergeEnabled();

            fc.Property(r => r.Name)
              .HasExcelIndex(1)
              .HasExcelTitle("Tên")
              .IsMergeEnabled();

            fc.Property(r => r.DisplayName)
              .HasExcelIndex(7)
              .HasExcelTitle("Tên đầy đủ")
              .IsIgnored(exportingIsIgnored: false, importingIsIgnored: true);

            fc.Property(r => r.Birthday)
              .HasExcelIndex(2)
              .HasExcelTitle("Ngày sinh")
              .HasDataFormatter("dd/MM/yyyy");


            fc.Property(r => r.Address)
              .HasExcelIndex(3)
              .HasExcelTitle("Địa chỉ");

            fc.Property(r => r.Age)
              .HasExcelIndex(4)
              .HasExcelTitle("Tuổi");

            fc.Property(r => r.Latitude)
              .HasExcelIndex(5)
              .HasExcelTitle("Vĩ độ");

            fc.Property(r => r.Longitude)
              .HasExcelIndex(6)
              .HasExcelTitle("Kinh Độ");
        }
    }
}
