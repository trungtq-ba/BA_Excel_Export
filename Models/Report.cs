using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations;

namespace BAExcelExport.Models
{
    public class Report
    {
        [Display(Name = "Thành phố")]
        public string City { get; set; }
        [Display(Name = "Tòa nhà")]
        public string Building { get; set; }
        [Display(Name = "Khu vực")]
        public string Area { get; set; }
        [Display(Name = "Thòi gian mở cửa")]
        public DateTime HandleTime { get; set; }
        [Display(Name = "Nhà môi giới")]
        public string Broker { get; set; }
        [Display(Name = "Khách hàng")]
        public string Customer { get; set; }
        [Display(Name = "Phòng")]
        public string Room { get; set; }
        [Display(Name = "Môi giới")]
        public decimal Brokerage { get; set; }
        [Display(Name = "Lợi nhuận")]
        public decimal Profits { get; set; }
    }
}
