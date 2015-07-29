using DistanceCalculator.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace DistanceCalculator.DTOs
{
    public class Response
    {
        public bool IsSucceed { get; set; }
        public string Message { get; set; }
        public string CalculatedAddressesFileName { get; set; }
        public ICollection<CalculatedMsa> CalculatedMsas { get; set; }
    }
}