using DistanceCalculator.DTOs;
using DistanceCalculator.Models;
using DistanceCalculator.ViewModels;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using System.Web.Script.Serialization;
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
        public string Test()
        {
            var Response = new Response();

            //try
            {
                if (Request.Files.Count > 0)
                {
                    var file = Request.Files[0];
                    if (file != null && file.ContentLength > 0)
                    {
                        var fileName = Path.GetFileName(file.FileName);
                        var FileNameWithoutExt = Path.GetFileNameWithoutExtension(file.FileName);
                        var FileExtension = Path.GetExtension(file.FileName);
                        
                        // 1st way to generate random string
                        var chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
                        var random = new Random();
                        var RandomString = new string(
                            Enumerable.Repeat(chars, 8)
                                      .Select(s => s[random.Next(s.Length)])
                                      .ToArray());

                        // second way to generate random UNIQUE string
                        Guid g = Guid.NewGuid();
                        string GuidString = Convert.ToBase64String(g.ToByteArray());
                        GuidString = GuidString.Replace("=", string.Empty);
                        GuidString = GuidString.Replace("+", string.Empty);
                        GuidString = GuidString.Replace("\\", string.Empty);
                        GuidString = GuidString.Replace("/", string.Empty);
                        
                        FileNameWithoutExt += GuidString;
                        fileName = FileNameWithoutExt + FileExtension;

                        var path = Path.Combine(Server.MapPath("~/XlsFiles/"), fileName);
                        file.SaveAs(path);
                        Microsoft.Office.Interop.Excel.Workbook Workbook = ExcelApp.Workbooks.Open(path);
                        
                        return new JavaScriptSerializer().Serialize(ProcessEachWorksheet(Workbook));
                        /////////////////////////////////////////////////////////////////////////////
                        /*
                        //Declare variables to hold refernces to Excel objects.
                        DocumentFormat.OpenXml.Spreadsheet.Workbook workBook;
                        SharedStringTable sharedStrings;
                        IEnumerable<Sheet> workSheets;
                        WorksheetPart custSheet;
                        WorksheetPart orderSheet;

                        //Declare helper variables.
                        string custID;
                        string orderID;

                        //Open the Excel workbook.
                        using (SpreadsheetDocument document =
                        SpreadsheetDocument.Open(path, true))
                        {
                            //References to the workbook and Shared String Table.
                            workBook = document.WorkbookPart.Workbook;
                            workSheets = workBook.Descendants<Sheet>();
                            sharedStrings = document.WorkbookPart.SharedStringTablePart.SharedStringTable;

                            //Reference to Target MSAs Excel Worksheet
                            var WorksheetId = workSheets.First(s => s.Name == @"Target MSAs").Id;
                            var WorksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(WorksheetId);

                            //LINQ query to skip first tow rows with column names.
                            IEnumerable<Row> dataRows =
                                from row in WorksheetPart.Worksheet.Descendants<Row>()
                                where row.RowIndex > 1
                                select row;
                            
                            List<MsaAddress> MsaAddresses = new List<MsaAddress>();
                            foreach (Row row in dataRows)
                            {
                                //LINQ query to return the row's cell values.
                                //Where clause filters out any cells that do not contain a value.
                                //Select returns the value of a cell unless the cell contains
                                //  a Shared String.
                                //If the cell contains a Shared String, its value will be a 
                                //  reference id which will be used to look up the value in the 
                                //  Shared String table.
                                IEnumerable<String> textValues =
                                    from cell in row.Descendants<Cell>()
                                    //where cell.CellValue != null
                                    select cell.CellValue.InnerText;
                                        //(//cell.DataType != null
                                            //cell.DataType.HasValue
                                            //cell.DataType == CellValues.SharedString
                                            //? sharedStrings.ChildElements[
                                            //int.Parse(cell.CellValue.InnerText)].InnerText
                                            //: cell.CellValue.InnerText);
                                
                                // convert IEnum to List
                                var AddressInfos = textValues.ToList();
                                
                                MsaAddress MsaAddress = new MsaAddress();
                                MsaAddress.Address = ((AddressInfos[0] != null) ? AddressInfos[0] : string.Empty);
                                MsaAddress.City = ((AddressInfos[1] != null) ? AddressInfos[1] : string.Empty);
                                MsaAddress.State = ((AddressInfos[2] != null) ? AddressInfos[2] : string.Empty);
                                MsaAddress.Zip = ((AddressInfos[3] != null) ? AddressInfos[3] : string.Empty);
                                MsaAddress.Phone = ((AddressInfos[4] != null) ? AddressInfos[4] : string.Empty);
                                MsaAddress.Center = ((AddressInfos[5] != null) ? AddressInfos[5] : string.Empty);
                                MsaAddress.CityState = ((AddressInfos[6] != null) ? AddressInfos[6] : string.Empty);
                                MsaAddress.MSA = ((AddressInfos[7] != null) ? AddressInfos[7] : string.Empty);

                                MsaAddresses.Add(MsaAddress);
                            }
                            
                            var AvailableMSAs = MsaAddresses.Where(msa => msa.MSA != null).Select(msa => msa.MSA).Distinct();
                        }
                        */

                        /*
                        string pathToExcelFile = path;
                        //string sheetName = "Target MSAs";
                        var excelFile = new ExcelQueryFactory(path);
                        var sheetNames = excelFile.GetWorksheetNames() as List<string>;
                        var artistAlbums = from a in excelFile.Worksheet(sheetNames[0]) where int.Parse(a.Zip()+"") > 30000 select a;
                        
                        excelFile.Worksheet(sheetNames[0]);
                        
                        foreach (var a in artistAlbums)
                        {
                            string artistInfo = "Artist Name: {0}; Album: {1}";
                            Console.WriteLine(string.Format(artistInfo, a["Name"], a["Title"]));
                        }
                        */
                    }
                }
            }
            //catch(Exception exc)
            {
                Response.IsSucceed = false;
                //Response.Message = exc.Message;
            }

            return new JavaScriptSerializer().Serialize(Response);
        }

        private Response ProcessEachWorksheet(Microsoft.Office.Interop.Excel.Workbook Workbook)
        {

            //try
            {
                int NumberOfSheets = Workbook.Sheets.Count;
                var Response = new Response();
                ICollection<CalculatedMsa> CalculatedMsas = new List<CalculatedMsa>();

                // loop through all worksheets of the browsed workbook
                for (int sheetNumber = 1; sheetNumber < NumberOfSheets + 1; sheetNumber++)
                {
                    Microsoft.Office.Interop.Excel.Worksheet workSheet = (Microsoft.Office.Interop.Excel.Worksheet)Workbook.Sheets[sheetNumber];
                    Range range = workSheet.UsedRange;
                    int rows_count = range.Rows.Count;

                    var MsaAddresses = new List<MsaAddress>();

                    // loop through all rows of worksheet. Start from row 1, first row is for headings
                    for (int row = 2; row <= rows_count; row++)
                    {
                        MsaAddress MsaAddress = new MsaAddress();
                        MsaAddress.Address = ((workSheet.Cells[row, 1].value != null) ? workSheet.Cells[row, 1].value.ToString() : null);
                        MsaAddress.City = ((workSheet.Cells[row, 2].value != null) ? workSheet.Cells[row, 2].value.ToString() : null);
                        MsaAddress.State = ((workSheet.Cells[row, 3].value != null) ? workSheet.Cells[row, 3].value.ToString() : null);
                        MsaAddress.Zip = ((workSheet.Cells[row, 4].value != null) ? workSheet.Cells[row, 4].value.ToString() : null);
                        MsaAddress.Phone = ((workSheet.Cells[row, 5].value != null) ? workSheet.Cells[row, 5].value.ToString() : null);
                        MsaAddress.Center = ((workSheet.Cells[row, 6].value != null) ? workSheet.Cells[row, 6].value.ToString() : null);
                        MsaAddress.CityState = ((workSheet.Cells[row, 7].value != null) ? workSheet.Cells[row, 7].value.ToString() : null);
                        MsaAddress.MSA = ((workSheet.Cells[row, 8].value != null) ? workSheet.Cells[row, 8].value.ToString() : null);

                        MsaAddresses.Add(MsaAddress);
                    }

                    var AvailableMSAsNames = MsaAddresses.Where(msa => msa.MSA != null).Select(msa => msa.MSA).Distinct();
                    
                    List<string> APIs = new List<string>
                    {
                        // sohailx2x
                        /*1*/ "AIzaSyDkajU8Ev-rg35iWxUBUFJOs10N9V36SaI",
                        /*2*/ "AIzaSyC4zkClQIrMMwl5X1brWH1sTW56MMNwxfs",
                        /*3*/ "AIzaSyBDovvzPb_5TEbiU7aYNvcr1h4eKI3OSxQ",
                        /*4*/ "AIzaSyCvwJMKAiyburZ5XhgqlvhHhVAH92APudU",
                        /*5*/ "AIzaSyARhou4fBIRdlVZZUGhLJO5v6mkwSq7hAo",
                        /*6*/ "AIzaSyBPxi8tCISPifCBA8n6XK-PwgO2qhFEA7I",
                        /*7*/ "AIzaSyBofjKVhrpkzy4BkDC9hE_MuBBhNTB6K7I",
                        /*8*/ "AIzaSyAkqd_iuOlcj5oXYiRI-fTdvRlt_nsCw2U",
                        /*9*/ "AIzaSyC1O843081Q6CoXlpiiCbKfZUkHDUs0C2c",
                        /*10*/ "AIzaSyB926nJW35_jdyKTbNNMMRKiMQBsAPJgSo",
                        /*11*/ "AIzaSyBoZ_EB87v7EY8kKP_EhQgzvHEN2llfAzI",
                        /*12*/ "AIzaSyAzSig0Wkp6CPXOJgw_tHWaXs5IuYVtJ4o",
                        /*13*/ "AIzaSyBRYokqwDFhG2ir4Gei-2t-VwP1IY21ynE",
                        /*14*/ "AIzaSyD51v-nkDoNp56QqlJL4FcfdMaAInxJ3r0",
                        /*15*/ "AIzaSyBLoywKIH9ImQ92l8s9nX9-IJLPcvXLwZg",
                        /*16*/ "AIzaSyBLCx46UHlw5nmEDmGCb4ZT9yt4Tm9EVGY",

                        // geemmii
                        /*1*/ "AIzaSyDq5hn0F2ewiaFxkmvaNKCIAhQyPhMbG8U",
                        /*2*/ "AIzaSyAiwyhxLoJxukF5n51KsMQ3d8_JZhHEWDY",
                        /*3*/ "AIzaSyDV48WQ7SqCNQaDlRdXCZrLkoCgxWeu1fs",
                        /*4*/ "AIzaSyDhqZrHoLBlUKesUQdm2tnYjLF4qKHruPg",
                        /*5*/ "AIzaSyAumu1SWpPP5ntjHMfHwRT4HjHuRyDLe9M",
                        /*6*/ "AIzaSyA6h60NvsG1KggY3Yf73ldDp4JiE9V64k0",
                        /*7*/ "AIzaSyABQ_1iR3ydAUNSs-TJ6isxsLhGNoiw35U",
                        /*8*/ "AIzaSyD4pPGjeZ750UPPpUVIuHS6dDRFcxv_r48",
                        /*9*/ "AIzaSyDTP-D6XCsuS4cI_7bA4C-BZl4sg2UER3k",
                        /*10*/ "AIzaSyBqCiJx3-w672wlnzzJjB1TyFmhBiJBJH0",
                    };
                    
                    foreach (var AvailableMsaName in AvailableMSAsNames)
                    {
                        CalculatedMsa CalculatedMsa = new CalculatedMsa();
                        CalculatedMsa.Name = AvailableMsaName;
                        CalculatedMsa.AddressesDistances = new List<AddressesDistance>();
                        string MatrixApiMode = "driving";
                        string MatrixApiLanguage = "en-EN";
                        
                        var SubMsaAddresses = MsaAddresses.Where(msa => msa.MSA.Equals(AvailableMsaName)).ToList<MsaAddress>();
                        for (int i = 0; i < SubMsaAddresses.Count(); i++)
                        {
                            string OriginDestinationsStr = "origins=" + HttpUtility.UrlEncode(SubMsaAddresses[i].Address) + "+" + HttpUtility.UrlEncode(SubMsaAddresses[i].City) + "+" + HttpUtility.UrlEncode(SubMsaAddresses[i].State) + "&destinations=";
                            int counter = 0;
                            int j = i + 1;

                            // if we reach last row as origin address
                            if (i < SubMsaAddresses.Count() && j == SubMsaAddresses.Count()) break;
                            
                            for ( ; j < SubMsaAddresses.Count(); j++)
                            {
                                if (counter++ == 0)
                                {
                                    OriginDestinationsStr += HttpUtility.UrlEncode(SubMsaAddresses[j].Address) + "+" + HttpUtility.UrlEncode(SubMsaAddresses[j].City) + "+" + HttpUtility.UrlEncode(SubMsaAddresses[j].State);
                                }
                                else
                                {
                                    OriginDestinationsStr += "|" + HttpUtility.UrlEncode(SubMsaAddresses[j].Address) + "+" + HttpUtility.UrlEncode(SubMsaAddresses[j].City) + "+" + HttpUtility.UrlEncode(SubMsaAddresses[j].State);
                                }
                            }

                            int RandomNumber = new Random().Next(0, APIs.Count - 1);
                            JavaScriptSerializer serializer = new JavaScriptSerializer();
                            string serviceUrl = string.Format("https://maps.googleapis.com/maps/api/distancematrix/json?" + OriginDestinationsStr + "&mode=" + MatrixApiMode + "&language=" + MatrixApiLanguage + "&key=" + APIs[RandomNumber]);

                            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(serviceUrl);
                            request.Method = "GET";
                            request.Accept = "application/json; charset=UTF-8";
                            request.Headers.Add("Accept-Language", " en-US");

                            try
                            {
                                var httpResponse = (HttpWebResponse)request.GetResponse();
                                DistanceMatrixResponse DistanceMatrixResponse = null;
                                using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                                {
                                    var responseText = streamReader.ReadToEnd();
                                    DistanceMatrixResponse = serializer.Deserialize<DistanceMatrixResponse>(responseText);
                                    if (DistanceMatrixResponse.Status == "OK")
                                    {
                                        if (DistanceMatrixResponse.Error_Message != null)
                                        {
                                            Response.IsSucceed = false;
                                            Response.Message = DistanceMatrixResponse.Error_Message;
                                            return Response;
                                        }

                                        var Elements = DistanceMatrixResponse.Rows.First().Elements.ToList();
                                        var OriginAddress = DistanceMatrixResponse.Origin_Addresses.ToList().First();
                                        int ii = 0;
                                        foreach (var DestinationAddress in DistanceMatrixResponse.Destination_Addresses)
                                        {
                                            AddressesDistance AddressesDistance = new AddressesDistance();
                                            AddressesDistance.OriginAddress = OriginAddress;
                                            AddressesDistance.DestinationAddress = DestinationAddress;

                                            if (Elements[ii].Status.Equals("OK"))
                                            {
                                                AddressesDistance.Distance = Elements[ii].Distance.Text;
                                            }
                                            else
                                            {
                                                AddressesDistance.Distance = "No results found";
                                            }

                                            CalculatedMsa.AddressesDistances.Add(AddressesDistance);
                                            ii++;
                                        }
                                    }
                                    else
                                    {
                                        AddressesDistance AddressesDistance = new AddressesDistance();
                                        //AddressesDistance.OriginAddress = OriginAddress;
                                        //AddressesDistance.DestinationAddress = DestinationAddress;
                                        AddressesDistance.Distance = "Invalid Request";
                                        CalculatedMsa.AddressesDistances.Add(AddressesDistance);
                                    }
                                }
                            }
                            catch (WebException WebExc)
                            {
                                if (WebExc.Status == WebExceptionStatus.ProtocolError && WebExc.Response != null)
                                {
                                    j = i + 1;
                                    for (; j < SubMsaAddresses.Count(); j++)
                                    {
                                        OriginDestinationsStr = "origins=" + HttpUtility.UrlEncode(SubMsaAddresses[i].Address) + "+" + HttpUtility.UrlEncode(SubMsaAddresses[i].City) + "+" + HttpUtility.UrlEncode(SubMsaAddresses[i].State) + "&destinations=" + HttpUtility.UrlEncode(SubMsaAddresses[j].Address) + "+" + HttpUtility.UrlEncode(SubMsaAddresses[j].City) + "+" + HttpUtility.UrlEncode(SubMsaAddresses[j].State);
                                        serviceUrl = string.Format("https://maps.googleapis.com/maps/api/distancematrix/json?" + OriginDestinationsStr + "&mode=" + MatrixApiMode + "&language=" + MatrixApiLanguage + "&key=" + APIs[RandomNumber]);

                                        request = (HttpWebRequest)WebRequest.Create(serviceUrl);
                                        request.Method = "GET";
                                        request.Accept = "application/json; charset=UTF-8";
                                        request.Headers.Add("Accept-Language", " en-US");

                                        var httpResponse = (HttpWebResponse)request.GetResponse();
                                        DistanceMatrixResponse DistanceMatrixResponse = null;
                                        using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                                        {
                                            var responseText = streamReader.ReadToEnd();
                                            DistanceMatrixResponse = serializer.Deserialize<DistanceMatrixResponse>(responseText);
                                            if (DistanceMatrixResponse.Status == "OK")
                                            {
                                                if (DistanceMatrixResponse.Error_Message != null)
                                                {
                                                    Response.IsSucceed = false;
                                                    Response.Message = DistanceMatrixResponse.Error_Message;
                                                    return Response;
                                                }

                                                var Elements = DistanceMatrixResponse.Rows.First().Elements.ToList();
                                                var OriginAddress = DistanceMatrixResponse.Origin_Addresses.ToList().First();
                                                int ii = 0;
                                                foreach (var DestinationAddress in DistanceMatrixResponse.Destination_Addresses)
                                                {
                                                    AddressesDistance AddressesDistance = new AddressesDistance();
                                                    AddressesDistance.OriginAddress = OriginAddress;
                                                    AddressesDistance.DestinationAddress = DestinationAddress;

                                                    if (Elements[ii].Status.Equals("OK"))
                                                    {
                                                        AddressesDistance.Distance = Elements[ii].Distance.Text;
                                                    }
                                                    else
                                                    {
                                                        AddressesDistance.Distance = "No results found";
                                                    }

                                                    CalculatedMsa.AddressesDistances.Add(AddressesDistance);
                                                    ii++;
                                                }
                                            }
                                            else
                                            {
                                                AddressesDistance AddressesDistance = new AddressesDistance();
                                                //AddressesDistance.OriginAddress = OriginAddress;
                                                //AddressesDistance.DestinationAddress = DestinationAddress;
                                                AddressesDistance.Distance = "Invalid Request";
                                                CalculatedMsa.AddressesDistances.Add(AddressesDistance);
                                            }
                                        }
                                    }
                                }
                            }
                            catch (Exception Exc)
                            { }
                        }

                        CalculatedMsas.Add(CalculatedMsa);
                    }
                }
                
                Response.IsSucceed = true;
                Response.Message = "Addresses calculated.";
                Response.CalculatedMsas = CalculatedMsas;

                Microsoft.Office.Interop.Excel.Application myExcelFile = new Microsoft.Office.Interop.Excel.Application();
                Workbook myWorkBook = myExcelFile.Workbooks.Add(XlSheetType.xlWorksheet);
                Worksheet myWorkSheet = (Worksheet)myExcelFile.ActiveSheet;

                //excel sheet is a 1-based array
                myWorkSheet.Cells[1, 1] = "MSA";
                myWorkSheet.Cells[1, 2] = "Origin Address";
                myWorkSheet.Cells[1, 3] = "Destination Address";
                myWorkSheet.Cells[1, 4] = "Distance";

                // don't open excel file in windows during building
                myExcelFile.Visible = true;

                myWorkSheet.EnableAutoFilter = true;
                myWorkSheet.Cells.AutoFilter(1);

                //Set the header-row bold
                myWorkSheet.Range["A1", "A1"].EntireRow.Font.Bold = true;

                //Adjust all columns
                myWorkSheet.Columns.AutoFit();

                // since, first row has titles that are set above, start from row-2 and fill each row of excel file.
                int r = 2;
                foreach (var CalculatedMsa in CalculatedMsas)
                {
                    int c = 1;
                    foreach (var AddressesDistance in CalculatedMsa.AddressesDistances)
                    {
                        myWorkSheet.Cells[r, c] = CalculatedMsa.Name;
                        myWorkSheet.Cells[r, c + 1] = AddressesDistance.OriginAddress;
                        myWorkSheet.Cells[r, c + 2] = AddressesDistance.DestinationAddress;
                        myWorkSheet.Cells[r, c + 3] = AddressesDistance.Distance;
                        r++;
                    }
                }

                // set the font style of first row as Bold which has titles of each column
                myWorkSheet.Rows[1].Font.Bold = true;
                myWorkSheet.Rows[1].Font.Size = 12;

                // after filling, save the file to the specified location
                myWorkBook.SaveCopyAs(Path.Combine(Server.MapPath("~/XlsFiles/CalculatedAddresses.xlsx")));

                return Response;
            }
            //catch(Exception exc)
            {
                var Response = new Response();
                Response.IsSucceed = false;
            //    Response.Message = exc.Message;
                return Response;
            }
        }
    }
}