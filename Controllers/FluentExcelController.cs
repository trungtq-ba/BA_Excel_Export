using BAExcelExport.Extensions;
using BAExcelExport.Models;
using FluentExcel;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
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
        public ActionResult FluentExcelExport()
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();

            // global call this
            FluentConfiguration();

            // demo the extension point
            Excel.Setting.For<Report>().FromAnnotations()
                                       .AdjustAutoIndex();

            // Change title cell style
            //Excel.Setting.TitleCellStyleApplier = MyTitleCellApplier;

            var reports = new Report[MAX_RECORD];
            for (int i = 0; i < MAX_RECORD; i++)
            {
                reports[i] = new Report
                {
                    City = "ningbo",
                    Building = "世茂首府",
                    HandleTime = DateTime.Now.AddDays(7 * i),
                    Broker = "rigofunc 18957139**7",
                    Customer = "yingting 18957139**7",
                    Room = "103",
                    Brokerage = 125 * i,
                    Profits = 25 * i
                };
            }

           
            HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new ByteArrayContent(null)
            };

            response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

            string fileName = ReportHelper.GetFileName(typeof(Report).Name);

            sw.Stop();

            return File(response.Content.ReadAsByteArrayAsync().Result, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", sw.ElapsedMilliseconds + "_" + fileName);

        }

        /// <summary>
        /// Use fluent configuration api. (doesn't poison your POCO)
        /// </summary>
        private static void FluentConfiguration()
        {
            var fc = Excel.Setting.For<Report>();

            fc.HasStatistics("合计", "SUM", 6, 7);
              
            fc.Property(r => r.City)
              .HasExcelIndex(0)
              .HasExcelTitle("城市")
              .IsMergeEnabled()
              .HasDataFormatter("");
               

            // or
            //fc.Property(r => r.City).HasExcelCell(0,"城市", allowMerge: true);

            fc.Property(r => r.Building)
              .HasExcelIndex(1)
              .HasExcelTitle("楼盘")
              .IsMergeEnabled();

            // configures the ignore when exporting or importing.
            fc.Property(r => r.Area)
              .HasExcelIndex(8)
              .HasExcelTitle("Area")
              .IsIgnored(exportingIsIgnored: false, importingIsIgnored: true);

            // or
            //fc.Property(r => r.Area).IsIgnored(8, "Area", formatter: null, exportingIsIgnored: false, importingIsIgnored: true);

            fc.Property(r => r.HandleTime)
              .HasExcelIndex(2)
              .HasExcelTitle("成交时间")
              .HasDataFormatter("yyyy-MM-dd");

            // or
            //fc.Property(r => r.HandleTime).HasExcelCell(2, "成交时间", formatter: "yyyy-MM-dd", allowMerge: false);
            // or
            //fc.Property(r => r.HandleTime).HasExcelCell(2, "成交时间", "yyyy-MM-dd");

            fc.Property(r => r.Broker)
              .HasExcelIndex(3)
              .HasExcelTitle("经纪人");

            fc.Property(r => r.Customer)
              .HasExcelIndex(4)
              .HasExcelTitle("客户");

            fc.Property(r => r.Room)
              .HasExcelIndex(5)
              .HasExcelTitle("房源");

            fc.Property(r => r.Brokerage)
              .HasExcelIndex(6)
              .HasDataFormatter("￥0.00")
              .HasExcelTitle("佣金(元)");

            fc.Property(r => r.Profits)
              .HasExcelIndex(7)
              .HasExcelTitle("收益(元)");
        }
    }
}
