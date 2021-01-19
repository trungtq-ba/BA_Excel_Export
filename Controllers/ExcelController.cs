using FluentExcel;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;


namespace BAExcelExport.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
        static ExcelController()
        {
            FluentConfiguration();
        }

        private const int MAX_RECORD = 100;

        // GET api/Excel/excelExport
        [HttpGet]
        [Route("excelExport")]
        public ActionResult ExcelExport()
        {
            var dataInput = DataHelper.GenerateData(MAX_RECORD);

            List<ColumnInfo> settings = new List<ColumnInfo>()
            {
                new ColumnInfo(){ ColumnName="OrderNumber",Caption="STT", Format="",Visible=true,Width=100},
                new ColumnInfo(){ ColumnName="Name",Caption="Tên", Format="",Visible=true,Width=100},
                new ColumnInfo(){ ColumnName="DisplayName",Caption="Tên hiển thị", Format="",Visible=true,Width=100},
                new ColumnInfo(){ ColumnName="Address",Caption="Địa chỉ", Format="",Visible=true,Width=100},
                new ColumnInfo(){ ColumnName="Age",Caption="Tuổi", Format="",Visible=true,Width=100},
                new ColumnInfo(){ ColumnName="Latitude",Caption="Kinh độ", Format="",Visible=true,Width=100},
                new ColumnInfo(){ ColumnName="Longitude",Caption="Vĩ độ", Format="",Visible=true,Width=100},
                new ColumnInfo(){ ColumnName="Birthday",Caption="Ngày sinh", Format="HH:mm:ss dd-MM-yyyy",Visible=true,Width=100}
            };

            ReportSourceTemplate<ReportDataModel> template = new ReportSourceTemplate<ReportDataModel>()
            {
                ReportTitle = "BÁO CÁO DANH SÁCH ĐIỂM",
                ReportSubtitleLevel1 = "TIÊU ĐỀ CON CỦA BÁO CÁO DANH SÁCH ĐIỂM",
                ReportSubtitleLevel2 = "MÔ TẢ CỦA BÁO CÁO DANH SÁCH ĐIỂM",
                FileName = typeof(ReportDataModel).Name,
                SettingColumns = settings,
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
            
            Excel.Setting.AutoSizeColumnsEnabled = true;

            fc.HasStatistics("Tổng", "SUM", 5,6,7);

            fc.Property(r => r.OrderNumber)
              .HasExcelIndex(0)
              .HasExcelTitle("STT")
              .IsMergeEnabled();

            fc.Property(r => r.Name)
              .HasExcelIndex(1)
              .HasExcelTitle("Tên")
              .IsMergeEnabled();

            fc.Property(r => r.DisplayName)
              .HasExcelIndex(2)
              .HasExcelTitle("Tên đầy đủ")
              .IsIgnored(exportingIsIgnored: false, importingIsIgnored: true);

            fc.Property(r => r.Birthday)
              .HasExcelIndex(3)
              .HasExcelTitle("Ngày sinh")
              .HasDataFormatter("dd/MM/yyyy");

            fc.Property(r => r.Address)
              .HasExcelIndex(4)
              .HasExcelTitle("Địa chỉ");

            fc.Property(r => r.Age)
              .HasExcelIndex(5)
              .HasExcelTitle("Tuổi");

            fc.Property(r => r.Latitude)
              .HasExcelIndex(6)
              .HasExcelTitle("Vĩ độ");

            fc.Property(r => r.Longitude)
              .HasExcelIndex(7)
              .HasExcelTitle("Kinh Độ");
        }
    }
}
