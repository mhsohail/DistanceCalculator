using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace DistanceCalculator.Models
{
    public class Element
    {
        public Distance Distance { get; set; }
        public Duration Duration { get; set; }
        public string Status { get; set; }
    }
}