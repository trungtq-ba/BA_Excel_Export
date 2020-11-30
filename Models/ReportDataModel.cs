using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace BAExcelExport.Models
{
    public class ReportDataModel
    {
        public int OrderNumber { get; set; }

        public string Name { get; set; }

        public string DisplayName { get; set; }

        public string Address { get; set; }

        public int Age { get; set; }

        public double Latitude { get; set; }

        public double Longitude { get; set; }
    }
}
