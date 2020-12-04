﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace BAExcelExport
{
    /// <summary>
    /// Template cung cấp dữ liệu cho báo cáo
    /// </summary>
    /// <Modified>
    /// Name     Date         Comments
    /// trungtq  30/11/2020   created
    /// </Modified>
    [Serializable]
    public class ReportSourceTemplate<TEntity> where TEntity: ReportDataModelBase
    {
        /// <summary>
        /// Default contructor
        /// </summary>
        public ReportSourceTemplate() { }

        /// <summary>
        /// Overloading contructor
        /// </summary>
        /// <param name="fromDate">Từ ngày</param>
        /// <param name="toDate">Đến ngày</param>
        /// <param name="vehiclePlates">Danh sách biển số</param>
        /// <param name="reportList">Danh sách báo cáo</param>
        public ReportSourceTemplate(DateTime fromDate, DateTime toDate, List<TEntity> reportList)
        {
            this.FromDate = fromDate;
            this.ToDate = toDate;
            this.ReportList = reportList;
        }

        /// <summary>
        /// Tiêu đề của báo cáo
        /// </summary>
        public string ReportTitle { get; set; } = "REPORT TITLE";

        /// <summary>
        /// Noi dung: Nội dung: Thời gian vi phạm trên {Minutes} phút
        /// </summary>
        public string ReportSubtitleLevel1 { get; set; } = "REPORT SUBTITLE LEVEL 1";

        /// <summary>
        /// Ngày báo cáo
        /// </summary>
        public string ReportSubtitleLevel2 { get; set; } = "REPORT SUBTITLE LEVEL 1";

        /// <summary>
        /// Từ ngày
        /// </summary>
        public DateTime FromDate { get; set; }

        /// <summary>
        /// Đến ngày
        /// </summary>
        public DateTime ToDate { get; set; }

        public string FileName { get; set; }

        public string SheetName { get; set; }

        public string FileTemplatePath { get; set; }

        /// <summary>
        /// Dữ liệu cung cấp cho báo cáo
        /// </summary>
        public List<TEntity> ReportList { get; set; }

        public List<ColumnInfo> SettingColumns { get; set; }

    }
}
