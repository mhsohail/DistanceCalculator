using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace DistanceCalculator.Models
{
    public class DistanceMatrixResponse
    {
        public IEnumerable<string> Destination_Addresses { get; set; }
        public IEnumerable<string> Origin_Addresses { get; set; }
        public IEnumerable<Row> Rows { get; set; }
        public string Error_Message { get; set; }
        public string Status { get; set; }
    }
}