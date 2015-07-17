using DistanceCalculator.ViewModels;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using VdoValley.Attributes;

namespace DistanceCalculator.Controllers
{
    public class HomeController : Controller
    {
        Application ExcelApp;

        public HomeController()
        {
            ExcelApp = new Application();
        }

        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";
            return View();
        }

        //[AjaxRequestOnly]
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult GetDistances(HomeIndexViewModel model, HttpPostedFileBase File)
        {
            if(ModelState.IsValid)
            {
                if (Request.Files.Count > 0)
                {
                    var file = Request.Files[0];
                    if (file != null && file.ContentLength > 0)
                    {
                        var fileName = Path.GetFileName(file.FileName);
                        var path = Path.Combine(Server.MapPath("~/XlsFiles/"), fileName);
                        file.SaveAs(path);
                     }
                }

                //return RedirectToAction("UploadDocument");
            }
            
            /*
             * Reference: http://stackoverflow.com/questions/304617/html-helper-for-input-type-file
            */
            using (MemoryStream memoryStream = new MemoryStream())
            {
                //model.FilePath.InputStream.CopyTo(memoryStream);
            }
            
            //Workbook Workbook = ExcelApp.Workbooks.Open(FullFilePath);
            return View(model);
        }

        //[AjaxRequestOnly]
        [HttpPost]
        public void Test()
        {
            if (Request.Files.Count > 0)
            {
                var file = Request.Files[0];
                if (file != null && file.ContentLength > 0)
                {
                    var fileName = Path.GetFileName(file.FileName);
                    var path = Path.Combine(Server.MapPath("~/XlsFiles/"), fileName);
                    file.SaveAs(path);
                }
            }

            /*
             * Reference: http://stackoverflow.com/questions/304617/html-helper-for-input-type-file
            */
            using (MemoryStream memoryStream = new MemoryStream())
            {
                //model.FilePath.InputStream.CopyTo(memoryStream);
            }

            //Workbook Workbook = ExcelApp.Workbooks.Open(FullFilePath);
        }
    }
}