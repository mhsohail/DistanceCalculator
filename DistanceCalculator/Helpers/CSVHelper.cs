using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace DistanceCalculator.Helpers
{
    public static class CSVHelper
    {
        public static void ExportCSV(string csv, string filename)
        {
            HttpContext.Current.Response.Clear();
            HttpContext.Current.Response.AddHeader("content-disposition", string.Format("attachment; filename={0}.csv", filename));
            HttpContext.Current.Response.ContentType = "text/csv";
            HttpContext.Current.Response.AddHeader("Pragma", "public");
            HttpContext.Current.Response.Write(csv);
            HttpContext.Current.Response.End();
        }
    }
}