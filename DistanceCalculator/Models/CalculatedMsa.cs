using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace DistanceCalculator.Models
{
    public class CalculatedMsa
    {
        public string Name { get; set; }
        public ICollection<AddressesDistance> AddressesDistances { get; set; }
    }
}