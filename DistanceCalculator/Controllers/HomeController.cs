using DistanceCalculator.Classes;
using DistanceCalculator.DTOs;
using DistanceCalculator.Models;
using DistanceCalculator.ViewModels;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using System.Web.Script.Serialization;
using VdoValley.Attributes;
using System;
using System.IO;

namespace DistanceCalculator.Controllers
{
    public class HomeController : Controller
    {
        //Application ExcelApp;
        List<MsaAddress> MsaAddresses;

        public HomeController()
        {
            //ExcelApp = new Application();
            MsaAddresses = new List<MsaAddress>();
        }

        public ActionResult Index()
        {
            return View();
        }

        public void InsertText(SpreadsheetDocument spreadSheet, string text, string Row, uint Column)
        {    
            // Get the SharedStringTablePart. If it does not exist, create a new one.
            SharedStringTablePart shareStringPart;
            if (spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
            {
                shareStringPart = spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
            }
            else
            {
                shareStringPart = spreadSheet.WorkbookPart.AddNewPart<SharedStringTablePart>();
            }

            // Insert the text into the SharedStringTablePart.
            int index = InsertSharedStringItem(text, shareStringPart);

            // Insert a new worksheet.
            WorksheetPart worksheetPart = InsertWorksheet(spreadSheet.WorkbookPart);

            // Insert cell A1 into the new worksheet.
            Cell cell = InsertCellInWorksheet(Row, Column, worksheetPart);

            // Set the value of cell A1.
            cell.CellValue = new CellValue(index.ToString());
            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

            // Save the new worksheet.
            worksheetPart.Worksheet.Save();
        }

        private void CreateSpreadsheetWorkbook(string filepath)
        {
            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.
                Create(filepath, SpreadsheetDocumentType.Workbook);

            // Add a WorkbookPart to the document.
            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook.
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
                AppendChild<Sheets>(new Sheets());

            // Append a new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.
                    GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "mySheet"
            };
            sheets.Append(sheet);

            workbookpart.Workbook.Save();

            // Close the document.
            spreadsheetDocument.Close();
        }

        // Given a WorkbookPart, inserts a new worksheet.
        private WorksheetPart InsertWorksheet(WorkbookPart workbookPart)
        {
            // We need single sheet only, if there is a sheet, return
            if (workbookPart.WorksheetParts.Count() > 0)
            {
                return workbookPart.WorksheetParts.FirstOrDefault<WorksheetPart>();
            }

            // Add a new worksheet part to the workbook.
            WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());
            newWorksheetPart.Worksheet.Save();

            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

            // Get a unique ID for the new sheet.
            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Count() > 0)
            {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            string sheetName = "Sheet" + sheetId;

            // Append the new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
            sheets.Append(sheet);
            workbookPart.Workbook.Save();

            return newWorksheetPart;
        }

        // Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
        // If the cell already exists, returns it. 
        private Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            DocumentFormat.OpenXml.Spreadsheet.Row row;
            if (sheetData.Elements<DocumentFormat.OpenXml.Spreadsheet.Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<DocumentFormat.OpenXml.Spreadsheet.Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new DocumentFormat.OpenXml.Spreadsheet.Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // If there is not a cell with the specified column name, insert one.  
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }

        // Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
        // and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
        private int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            // If the part does not contain a SharedStringTable, create one.
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }

            int i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }

                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
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

        [AjaxRequestOnly]
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult GetDistances(HomeIndexViewModel model, HttpPostedFileBase File)
        {
            if (ModelState.IsValid)
            {
                if (Request.Files.Count > 0)
                {
                    var file = Request.Files[0];
                    if (file != null && file.ContentLength > 0)
                    {
                        var fileName = Path.GetFileName(file.FileName);
                        var path = Path.Combine(Server.MapPath("~/App_Data/"), fileName);
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

        [AjaxRequestOnly]
        public string CalculateAddresses(List<MsaAddress> MsaAddresses)
        {
            return new JavaScriptSerializer().Serialize(CalculateDistances(MsaAddresses));
        }

        public string ReadWrite()
        {
            // save results to excel file
            string FilePath = Path.Combine(Server.MapPath("~/App_Data/ReadWrite.xlsx"));
            bool fileExists = true;
            if (!System.IO.File.Exists(FilePath))
            {
                CreateSpreadsheetWorkbook(FilePath);
                fileExists = false;
            }
            
            uint i = 17;
            do
            {
                using (FileStream fs = new FileStream(FilePath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
                {
                    using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(fs, true))
                    {
                        InsertText(spreadSheetDocument, "MSA", Convert.ToChar(65 + 0).ToString(), i);
                        InsertText(spreadSheetDocument, "Origin Address", Convert.ToChar(65 + 1).ToString(), i);
                        InsertText(spreadSheetDocument, "Destination Address", Convert.ToChar(65 + 2).ToString(), i);
                        InsertText(spreadSheetDocument, "Distance", Convert.ToChar(65 + 3).ToString(), i);
                        
                        spreadSheetDocument.Close(); 
                    }
                }
                i++;
            }
            while(i < 21);

            return FilePath + " - " + fileExists;
        }

        [AjaxRequestOnly]
        public string PutResultsInExcel(List<CalculatedMsa> CalculatedMsas)
        {
            // create excel file to save results in
            DistanceCalculator.DTOs.Response Response = new DistanceCalculator.DTOs.Response();

            Guid g = Guid.NewGuid();
            string GuidString = Convert.ToBase64String(g.ToByteArray());
            GuidString = GuidString.Replace("=", string.Empty);
            GuidString = GuidString.Replace("+", string.Empty);
            GuidString = GuidString.Replace("\\", string.Empty);
            GuidString = GuidString.Replace("/", string.Empty);

            // save results to excel file
            string FilePath = Path.Combine(Server.MapPath("~/App_Data/CalculatedAddresses-" + GuidString + ".xlsx"));
            CreateSpreadsheetWorkbook(FilePath);
            using (FileStream fs = new FileStream(FilePath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
            {
                using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(fs, true))
                {
                    // put headings in first row
                    InsertText(spreadSheetDocument, "MSA", Convert.ToChar(65 + 0).ToString(), 1);
                    InsertText(spreadSheetDocument, "Origin Address", Convert.ToChar(65 + 1).ToString(), 1);
                    InsertText(spreadSheetDocument, "Destination Address", Convert.ToChar(65 + 2).ToString(), 1);
                    InsertText(spreadSheetDocument, "Distance", Convert.ToChar(65 + 3).ToString(), 1);

                    uint row = 2;
                    foreach (var CalculatedMsa in CalculatedMsas)
                    {
                        // if there is only one address for an MSA, this value will be null
                        if (CalculatedMsa.AddressesDistances != null)
                        {
                            foreach (var AddressesDistance in CalculatedMsa.AddressesDistances)
                            {
                                InsertText(spreadSheetDocument, CalculatedMsa.Name, Convert.ToChar(65 + 0).ToString(), row);
                                InsertText(spreadSheetDocument, AddressesDistance.OriginAddress, Convert.ToChar(65 + 1).ToString(), row);
                                InsertText(spreadSheetDocument, AddressesDistance.DestinationAddress, Convert.ToChar(65 + 2).ToString(), row);
                                InsertText(spreadSheetDocument, AddressesDistance.Distance, Convert.ToChar(65 + 3).ToString(), row);
                                row++;
                            }
                        }
                    }
                    spreadSheetDocument.Close();
                    Response.CalculatedAddressesFileName = "CalculatedAddresses-" + GuidString + ".xlsx";
                }
            }
            
            return new JavaScriptSerializer().Serialize(Response);
        }

        [AjaxRequestOnly]
        [HttpPost]
        public string Test()
        {
            var Response = new Response();

            try
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

                        var path = Path.Combine(Server.MapPath("~/App_Data/"), fileName);
                        file.SaveAs(path);

                        GetDataAndPutInModel(path);
                        return new JavaScriptSerializer().Serialize(MsaAddresses);
                        //return new JavaScriptSerializer().Serialize(CalculateDistances());
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
            catch(Exception exc)
            {
                Response.IsSucceed = false;
                Response.Message = exc.Message;
            }

            return new JavaScriptSerializer().Serialize(Response);
        }

        public ActionResult Download(string FileName)
        {
            //byte[] fileBytes = System.IO.File.ReadAllBytes("~/App_Data/CalculatedAddresses.xlsx");
            //var response = new FileContentResult(fileBytes, "application/octet-stream");
            //response.FileDownloadName = "~/App_Data/CalculatedAddresses.xlsx";
            //return response;

            string filepath = AppDomain.CurrentDomain.BaseDirectory + "/App_Data/" + FileName;
            byte[] filedata = System.IO.File.ReadAllBytes(filepath);
            string contentType = MimeMapping.GetMimeMapping(filepath);

            System.Net.Mime.ContentDisposition cd = new System.Net.Mime.ContentDisposition
            {
                FileName = FileName,
                Inline = true,
            };

            Response.AppendHeader("Content-Disposition", cd.ToString());

            return File(filedata, contentType);

        }

        // Retrieve the value of a cell, given a file name, sheet name, 
        // and address name.
        private string GetDataAndPutInModel(string fileName)
        {
            string value = null;
            using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fs, false))
                {
                    WorkbookPart workbookPart = doc.WorkbookPart;
                    SharedStringTablePart sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                    SharedStringTable sst = sstpart.SharedStringTable;

                    WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                    Worksheet sheet = worksheetPart.Worksheet;

                    var cells = sheet.Descendants<Cell>();
                    var rows = sheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Row>();

                    //Console.WriteLine("Row count = {0}", rows.LongCount());
                    //Console.WriteLine("Cell count = {0}", cells.LongCount());

                    // One way: go through each cell in the sheet
                    foreach (Cell cell in cells)
                    {
                        if ((cell.DataType != null) && (cell.DataType == CellValues.SharedString))
                        {
                            int ssid = int.Parse(cell.CellValue.Text);
                            string str = sst.ChildElements[ssid].InnerText;
                            //Console.WriteLine("Shared string {0}: {1}", ssid, str);
                        }
                        else if (cell.CellValue != null)
                        {
                            //Console.WriteLine("Cell contents: {0}", cell.CellValue.Text);
                        }
                    }

                    // Or... via each row
                    foreach (DocumentFormat.OpenXml.Spreadsheet.Row row in rows.Skip(1))
                    {
                        /*
                        foreach (Cell c in row.Elements<Cell>())
                        {
                            if ((c.DataType != null) && (c.DataType == CellValues.SharedString))
                            {
                                //Console.WriteLine("Shared string {0}: {1}", ssid, str);
                            }
                            else if (c.CellValue != null)
                            {
                                //Console.WriteLine("Cell contents: {0}", c.CellValue.Text);
                            }
                        }
                        */

                        MsaAddress MsaAddress = new MsaAddress();

                        var AddressCell = row.Elements<Cell>().ToList()[0];
                        if ((AddressCell.DataType != null) && (AddressCell.DataType == CellValues.SharedString))
                        {
                            int ssid = int.Parse(AddressCell.CellValue.Text);
                            string str = sst.ChildElements[ssid].InnerText;
                            MsaAddress.Address = str;
                        }

                        var CityCell = row.Elements<Cell>().ToList()[1];
                        if ((CityCell.DataType != null) && (CityCell.DataType == CellValues.SharedString))
                        {
                            int ssid = int.Parse(CityCell.CellValue.Text);
                            string str = sst.ChildElements[ssid].InnerText;
                            MsaAddress.City = str;
                        }

                        var StateCell = row.Elements<Cell>().ToList()[2];
                        if ((StateCell.DataType != null) && (StateCell.DataType == CellValues.SharedString))
                        {
                            int ssid = int.Parse(StateCell.CellValue.Text);
                            string str = sst.ChildElements[ssid].InnerText;
                            MsaAddress.State = str;
                        }

                        var ZipCell = row.Elements<Cell>().ToList()[3];
                        if ((ZipCell.DataType != null) && (ZipCell.DataType == CellValues.SharedString))
                        {
                            int ssid = int.Parse(ZipCell.CellValue.Text);
                            string str = sst.ChildElements[ssid].InnerText;
                            MsaAddress.Zip = str;
                        }

                        var PhoneCell = row.Elements<Cell>().ToList()[4];
                        if ((PhoneCell.DataType != null) && (PhoneCell.DataType == CellValues.SharedString))
                        {
                            int ssid = int.Parse(PhoneCell.CellValue.Text);
                            string str = sst.ChildElements[ssid].InnerText;
                            MsaAddress.Phone = str;
                        }

                        var CenterCell = row.Elements<Cell>().ToList()[5];
                        if ((CenterCell.DataType != null) && (CenterCell.DataType == CellValues.SharedString))
                        {
                            int ssid = int.Parse(CenterCell.CellValue.Text);
                            string str = sst.ChildElements[ssid].InnerText;
                            MsaAddress.Center = str;
                        }

                        var CityStateCell = row.Elements<Cell>().ToList()[6];
                        if ((CityStateCell.DataType != null) && (CityStateCell.DataType == CellValues.SharedString))
                        {
                            int ssid = int.Parse(CityStateCell.CellValue.Text);
                            string str = sst.ChildElements[ssid].InnerText;
                            MsaAddress.CityState = str;
                        }

                        var MSACell = row.Elements<Cell>().ToList()[7];
                        if ((MSACell.DataType != null) && (MSACell.DataType == CellValues.SharedString))
                        {
                            int ssid = int.Parse(MSACell.CellValue.Text);
                            string str = sst.ChildElements[ssid].InnerText;
                            MsaAddress.MSA = str;
                        }

                        MsaAddresses.Add(MsaAddress);
                    }
                }
            }

            return value;
        }

        private Response CalculateDistances(List<MsaAddress> MsaAddresses)
        {
            try
            {
                var Response = new Response();
                ICollection<CalculatedMsa> CalculatedMsas = new List<CalculatedMsa>();

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

                int ApiIndex = 0;
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

                        for (; j < SubMsaAddresses.Count(); j++)
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

                        JavaScriptSerializer serializer = new JavaScriptSerializer();
                        string serviceUrl = string.Format("https://maps.googleapis.com/maps/api/distancematrix/json?" + OriginDestinationsStr + "&mode=" + MatrixApiMode + "&language=" + MatrixApiLanguage + "&key=" + APIs[ApiIndex++%APIs.Count]);

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
                            try
                            {
                                if (WebExc.Status == WebExceptionStatus.ProtocolError && WebExc.Response != null)
                                {
                                    j = i + 1;
                                    for (; j < SubMsaAddresses.Count(); j++)
                                    {
                                        OriginDestinationsStr = "origins=" + HttpUtility.UrlEncode(SubMsaAddresses[i].Address) + "+" + HttpUtility.UrlEncode(SubMsaAddresses[i].City) + "+" + HttpUtility.UrlEncode(SubMsaAddresses[i].State) + "&destinations=" + HttpUtility.UrlEncode(SubMsaAddresses[j].Address) + "+" + HttpUtility.UrlEncode(SubMsaAddresses[j].City) + "+" + HttpUtility.UrlEncode(SubMsaAddresses[j].State);
                                        serviceUrl = string.Format("https://maps.googleapis.com/maps/api/distancematrix/json?" + OriginDestinationsStr + "&mode=" + MatrixApiMode + "&language=" + MatrixApiLanguage + "&key=" + APIs[ApiIndex++ % APIs.Count]);

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
                            catch(Exception exc)
                            {
                                Response.IsSucceed = false;
                                Response.Message = exc.Message;
                            }
                        }
                        catch (Exception Exc)
                        {
                            Response.IsSucceed = false;
                            Response.Message = Exc.Message; 
                        }
                    }

                    CalculatedMsas.Add(CalculatedMsa);
                }

                Response.IsSucceed = true;
                Response.Message = "Addresses calculated.";
                Response.CalculatedMsas = CalculatedMsas;

                // create excel file to save results in
                Guid g = Guid.NewGuid();
                string GuidString = Convert.ToBase64String(g.ToByteArray());
                GuidString = GuidString.Replace("=", string.Empty);
                GuidString = GuidString.Replace("+", string.Empty);
                GuidString = GuidString.Replace("\\", string.Empty);
                GuidString = GuidString.Replace("/", string.Empty);

                // save results to excel file
                string FilePath = Path.Combine(Server.MapPath("~/App_Data/CalculatedAddresses-" + GuidString + ".xlsx"));
                CreateSpreadsheetWorkbook(FilePath);
                using (FileStream fs = new FileStream(FilePath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
                {
                    using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(fs, true))
                    {
                        // put headings in first row
                        InsertText(spreadSheetDocument, "MSA", Convert.ToChar(65 + 0).ToString(), 1);
                        InsertText(spreadSheetDocument, "Origin Address", Convert.ToChar(65 + 1).ToString(), 1);
                        InsertText(spreadSheetDocument, "Destination Address", Convert.ToChar(65 + 2).ToString(), 1);
                        InsertText(spreadSheetDocument, "Distance", Convert.ToChar(65 + 3).ToString(), 1);

                        uint row = 2;
                        foreach (var CalculatedMsa in CalculatedMsas)
                        {
                            foreach (var AddressesDistance in CalculatedMsa.AddressesDistances)
                            {
                                InsertText(spreadSheetDocument, CalculatedMsa.Name, Convert.ToChar(65 + 0).ToString(), row);
                                InsertText(spreadSheetDocument, AddressesDistance.OriginAddress, Convert.ToChar(65 + 1).ToString(), row);
                                InsertText(spreadSheetDocument, AddressesDistance.DestinationAddress, Convert.ToChar(65 + 2).ToString(), row);
                                InsertText(spreadSheetDocument, AddressesDistance.Distance, Convert.ToChar(65 + 3).ToString(), row);
                                row++;
                            }
                        }
                        spreadSheetDocument.Close();
                        Response.CalculatedAddressesFileName = "CalculatedAddresses-" + GuidString + ".csv";
                    }
                }
                
                /*
                // save results to excel file
                try
                {
                    //CalculatedMsas.ToList<CalculatedMsa>().ExportCSV("myCSV");
                    Guid guid = Guid.NewGuid();
                    GuidString = Convert.ToBase64String(guid.ToByteArray());
                    GuidString = GuidString.Replace("=", string.Empty);
                    GuidString = GuidString.Replace("+", string.Empty);
                    GuidString = GuidString.Replace("\\", string.Empty);
                    GuidString = GuidString.Replace("/", string.Empty);
                    var list = CalculatedMsas.ToList<CalculatedMsa>()[0].AddressesDistances.ToList<AddressesDistance>();
                    Extensions.CreateCSVFromGenericList<AddressesDistance>(list, Path.Combine(Server.MapPath("~/App_Data/CalculatedAddresses-" + GuidString + ".csv")));
                    Response.CalculatedAddressesFileName = "CalculatedAddresses-" + GuidString + ".csv";
                }
                catch (System.Threading.ThreadAbortException exc)
                {
                    //Thrown then calling Response.End in CSV Helper
                    Response.IsSucceed = false;
                    Response.Message = exc.Message;
                }
                catch (Exception exc)
                {
                    Response.IsSucceed = false;
                    Response.Message = exc.Message;
                }
                */

                //        // set the font style of first row as Bold which has titles of each column
                //        myWorkSheet.Rows[1].Font.Bold = true;
                //        myWorkSheet.Rows[1].Font.Size = 12;

                //        // after filling, save the file to the specified location
                //        myWorkBook.SaveCopyAs(Path.Combine(Server.MapPath("~/App_Data/CalculatedAddresses.xlsx")));

                return Response;
            }
            catch (Exception exc)
            {
                var Response = new Response();
                Response.IsSucceed = false;
                Response.Message = exc.Message;
                return Response;
            }
        }
    }
}