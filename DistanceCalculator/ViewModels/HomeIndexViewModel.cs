using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace DistanceCalculator.ViewModels
{
    public class HomeIndexViewModel
    {
        [Key]
        public int Id { get; set; }

        public string ExcelFile { get; set; }

        [NotMapped]
        [Required]
        public HttpPostedFileBase File { get; set; }

        //[Required(ErrorMessage = "A header image is required"), FileExtensions(ErrorMessage = "Please upload an image file.")]
        //public string FileName
        //{
        //    get
        //    {
        //        if (File != null)
        //            return File.FileName;
        //        else
        //            return String.Empty;
        //    }
        //}
    }
}