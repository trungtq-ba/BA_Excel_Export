using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace BAExcelExport.ExcelExport
{
    public interface IAbstractDataExport
    {
        HttpResponseMessage Export<T>(List<T> exportData, string fileName, string sheetName);
    }


    public abstract class DataExportBase : IAbstractDataExport
    {
        protected string _sheetName;
        protected string _fileName;
        protected List<string> _headers;
        protected List<string> _type;
        protected IWorkbook _workbook;
        protected ISheet _sheet;
        private const string DefaultSheetName = "Sheet1";

        public string FileName
        {
            get
            {
                return $"{_fileName}_{DateTime.Now.ToString("yyyyMMddHHmmss")}.xlsx";
            }
        }

        public HttpResponseMessage Export<T>(List<T> exportData, string fileName, string sheetName = DefaultSheetName)
        {
            _fileName = fileName;
            _sheetName = sheetName;

            _workbook = new XSSFWorkbook();
            _sheet = _workbook.CreateSheet(_sheetName);

            var headerStyle = _workbook.CreateCellStyle();
            headerStyle.Alignment = HorizontalAlignment.Center;
            var headerFont = _workbook.CreateFont();
            headerFont.IsBold = true;
            headerStyle.SetFont(headerFont);

            WriteData(exportData);

            //Header
            var header = _sheet.CreateRow(0);
            for (var i = 0; i < _headers.Count; i++)
            {
                var cell = header.CreateCell(i);
                cell.SetCellValue(_headers[i]);
                cell.CellStyle = headerStyle;
                // It's heavy, it slows down your Excel if you have large data                
                _sheet.AutoSizeColumn(i);
            }

            using (var memoryStream = new MemoryStream())
            {
                _workbook.Write(memoryStream);
                var response = new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ByteArrayContent(memoryStream.ToArray())
                };

                response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

                return response;
            }
        }

        /// <summary>
        /// Generic Definition to handle all types of List
        /// </summary>
        /// <param name="exportData"></param>
        public abstract void WriteData<T>(List<T> exportData);
    }
}
