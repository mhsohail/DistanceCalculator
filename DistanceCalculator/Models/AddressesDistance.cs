using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace DistanceCalculator.Models
{
    public class AddressesDistance
    {
        public string MsaName { get; set; }
        public string OriginAddress { get; set; }
        public string DestinationAddress { get; set; }
        public string Distance { get; set; }
    }
}