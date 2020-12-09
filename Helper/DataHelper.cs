using BAExcelExport.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace BAExcelExport
{
    public class DataHelper
    {
        protected static string[] StreetNames = {
            "NGUYỄN CẢNH DỊ, PHƯỜNG ĐẠI KIM, HOÀNG MAI, HÀ NỘI.",
            "đường Ngô Xuân Quảng, Thị Trấn Trâu Quỳ, Huyện Gia Lâm, Thành phố Hà Nội",
            "đường Khuất Duy Tiến, Phường Thanh Xuân Bắc, Quận Thanh Xuân, Thành phố Hà Nội",
            "Xóm Phồ, Thọ Vực, Xã Đồng Tháp, Huyện Đan Phượng, Thành phố Hà Nội",
            "Thái Thịnh, Phường Thịnh Quang, Quận Đống Đa, Thành phố Hà Nội",
            "Giải Phóng, Phường Thịnh Liệt, Quận Hoàng Mai, Thành phố Hà Nội",
            "đường Thọ Vực, Cụm 8, Xã Thọ An, Huyện Đan Phượng, Thành phố Hà Nội",
            "Đại Từ, Phường Đại Kim, Quận Hoàng Mai, Thành phố Hà Nội",
            "Đường Louis 6 Khu Đô Thị Louis City Đại Mỗ, Phường Đại Mỗ, Quận Nam Từ Liêm, Thành phố Hà Nội",
            "Thôn Phù Lưu, Xã Phù Lưu Tế, Huyện Mỹ Đức, Thành phố Hà Nội",
            "ngách 163/1 ngõ 137 đường Đại Mỗ, Phường Đại Mỗ, Quận Nam Từ Liêm, Thành phố Hà Nội",
            "Tập Thể Binh Đoàn 12,Tổ 1, Phường Lĩnh Nam, Quận Hoàng Mai, Thành phố Hà Nội"
        };
        protected static string[] Names = {
            "TRUNGTQ",
            "HANHTH",
             "NAMTH",
            "LONGTQ",
             "QUOCNVC",
            "DONGHL",
             "LINHLV",
            "TRINHTX",
             "SONKT",
            "CONGND",
             "SONNL",
            "TRONGDC"
        };
        protected static string[] DisplayNames = {
            "TRẦN QUANG TRUNG",
            "TRẦN HỒNG HẠNH",
             "TRẦN HOÀNG NAM",
            "TRẦN QUANG LONG",
             "NGUYỄN VĂN CƯỜNG QUỐC",
            "LÊ HUY ĐÔNG",
             "LƯU VĂN LINH",
            "TRẦN XUÂN TRINH",
             "KHIẾU TRUNG SƠN",
            "NGUYỄN ĐỖ CÔNG",
             "NGUYỄN LUYỆN SƠN",
            "ĐINH CÔNG TRỌNG"
        };

        public static List<ReportDataModel> GenerateData(int maxRow)
        {
            Random rnd = new Random();

            return Enumerable.Range(1, maxRow).Select(index => new ReportDataModel()
            {
                OrderNumber = index,
                Name = $"{Names[rnd.Next(Names.Length)]}",
                DisplayName = $"{DisplayNames[rnd.Next(DisplayNames.Length)]}",
                Address = $"{rnd.Next(1, 100)} {StreetNames[rnd.Next(StreetNames.Length)]}",
                Age = index,
                Latitude = rnd.NextDouble(),
                Longitude = rnd.NextDouble(),
                Birthday = new DateTime(rnd.Next(1920, 2000), rnd.Next(1, 12), rnd.Next(1, 30))
            }).ToList();

        }
    }
}
