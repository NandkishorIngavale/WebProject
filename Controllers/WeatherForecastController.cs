using ExcelDataReader;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Syncfusion.Licensing;
using Syncfusion.Lic;
using Syncfusion.XlsIO;
using Microsoft.AspNetCore.Hosting;
using OfficeOpenXml;
using NPOI.XSSF.UserModel;
using GemBox.Spreadsheet;
using System.Data.OleDb;
using System.Runtime.Versioning;
using Aspose.Cells.Cloud;
using Aspose.Cells;

namespace WebProject.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class DataReadingController : ControllerBase
    {

        private readonly IWebHostEnvironment _env;
        private string sPath = "Bulk Records.xlsx";
        public DataReadingController(IWebHostEnvironment env)
        {
            _env = env;
        }
        [HttpGet]
        [Route("exceldatareader")]
        public IActionResult GetReadingFromExceldatareader(string filename)
        {
            try
            {
                IExcelDataReader reader = null;
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                string FilePath = _env.ContentRootPath + "\\" + "Bulk Records.xlsx";
                int content = 0;
                using var filestream = new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                if (Path.GetExtension(filename).Equals(".xls"))
                    reader = ExcelReaderFactory.CreateBinaryReader(filestream);
                else if (Path.GetExtension(filename).Equals(".xlsx"))
                    reader = ExcelReaderFactory.CreateOpenXmlReader(filestream);
                if (reader != null)
                {
                    content = reader.FieldCount;
                    while (reader.Read())
                    {
                        object i = reader.GetValue(0);
                    }
                    //var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                    //{
                    //    ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                    //    {
                    //        UseHeaderRow = true
                    //    }
                    //});
                    //https://discoverdot.net/projects/excel-data-reader
                    //DataTable s = result.Tables[0];
                }
                return Ok(content);
            }
            catch (Exception ex)
            {
                return Ok(ex.Message + " " + ex.StackTrace.ToString());
            }

        }
        //epplus and npo
        [HttpGet]
        [Route("syncfusion")]
        public IActionResult GetReadingFromSyncfusion(string filename)
        {
            try
            {
                ExcelEngine excelEngine = new ExcelEngine();
                // string FilePath = _env.ContentRootPath + "\\" + "Bulk Records.xlsx";
                using (var stream = new FileStream(filename, FileMode.Open, FileAccess.Read))
                {
                    var workbook = excelEngine.Excel.Workbooks.Open(stream);
                    return Ok(workbook.MaxRowCount);
                }
            }
            catch (Exception ex)
            {

                return Ok(ex.Message + " " + ex.StackTrace.ToString());
            }

        }
        [HttpGet]
        [Route("epplus")]
        public IActionResult GetReadingFromEPPlus(string filename)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.Commercial;
                using (var pck = new OfficeOpenXml.ExcelPackage())
                {
                    using (var stream = new FileStream(filename, FileMode.Open, FileAccess.Read))
                    {

                        pck.Load(stream);
                    }
                    var ws = pck.Workbook.Worksheets.First();
                    return Ok(ws.Workbook.Worksheets.Count);
                }
            }
            catch (Exception ex)
            {

                return Ok(ex.Message + " " + ex.StackTrace.ToString());
            }

        }
        [HttpGet]
        [Route("npoi")]
        public IActionResult GetReadingFromnpoi(string filename)
        {
            try
            {

                using (var stream = new FileStream(filename, FileMode.Open, FileAccess.Read))
                {
                    XSSFWorkbook xssWorkbook = new XSSFWorkbook(stream);
                    return Ok(xssWorkbook.Count);
                }
            }
            catch (Exception ex)
            {
                return Ok(ex.Message + " " + ex.StackTrace.ToString());
            }
        }
        [HttpGet]
        [Route("gemboxspreadsheet")]
        public IActionResult GetReadingFromGemBoxSpreadsheet(string filename)
        {
            try
            {

                SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
                // Load Excel workbook from file's path.
                ExcelFile workbook = ExcelFile.Load(filename);
                return Ok();
            }
            catch (Exception ex)
            {
                return Ok(ex.Message + " " + ex.StackTrace.ToString());
            }
        }

        [HttpGet]
        [Route("getexcel")]
        public IActionResult GetExcelFromDataTable()
        {
            try
            {
                //create a excel file
                
                //Instantiating a Workbook object
                Workbook workbook = new Workbook();

                //Adding a new worksheet to the Workbook object
                int i = workbook.Worksheets.Add();

                //Obtaining the reference of the newly added worksheet by passing its sheet index
                Worksheet worksheet = workbook.Worksheets[i];
                //Instantiating a "Products" DataTable object
                DataTable dataTable = new DataTable("Products");

                //Adding columns to the DataTable object
                dataTable.Columns.Add("Product ID", typeof(Int32));
                dataTable.Columns.Add("Product Name", typeof(string));
                dataTable.Columns.Add("Units In Stock", typeof(Int32));

                //Creating an empty row in the DataTable object
                DataRow dr = dataTable.NewRow();

                //Adding data to the row
                dr[0] = 1;
                dr[1] = "Aniseed Syrup";
                dr[2] = 15;

                //Adding filled row to the DataTable object
                dataTable.Rows.Add(dr);

                //Creating another empty row in the DataTable object
                dr = dataTable.NewRow();

                //Adding data to the row
                dr[0] = 2;
                dr[1] = "Boston Crab Meat";
                dr[2] = 123;

                //Adding filled row to the DataTable object
                dataTable.Rows.Add(dr);

                //Importing the contents of DataTable to the worksheet starting from "A1" cell,
                //where true specifies that the column names of the DataTable would be added to
                //the worksheet as a header row
                worksheet.Cells.ImportDataTable(dataTable, true, "A1");
                var fileName = DateTime.Now.ToString("yyyyMMddhhmmss");
                workbook.Save(fileName + ".xls");


                Workbook workbookR = new Workbook(fileName + ".xls");

                // Using the Sheet 1 in Workbook
                Worksheet worksheetR = workbook.Worksheets[0];

                // Accessing a cell using its name
                Cell cell = worksheet.Cells["A1"];

                string value = cell.Value.ToString();


                //read an excel file

                return Ok(value);
            }
            catch (Exception ex)
            {
                return Ok(ex.Message + " " + ex.StackTrace.ToString());
            }
        }

        [HttpGet]
        [Route("oledb")]
        public IActionResult GetReadingFromOLEDB(string filename)
        {
            try
            {
                string strFilePath = filename;
                string connString = string.Empty;

                if (Path.GetExtension(strFilePath).ToLower().Trim() == ".xls" && Environment.Is64BitOperatingSystem == false)
                    connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strFilePath + ";Extended Properties=Excel 8.0;HDR=Yes;IMEX=2";
                else
                connString = "Provider=Microsoft.ACE.OLEDB.12.0; ";
                connString = connString + "Data Source='" + strFilePath;
                connString = connString + "';Extended Properties=\"Excel 12.0;HDR=YES;\"";
                DataSet excelDataSet = new DataSet();
                using (OleDbConnection conn = new OleDbConnection(connString))
                {
                    conn.Open();
                    OleDbDataAdapter objDA = new System.Data.OleDb.OleDbDataAdapter
                    ("select * from [Sheet1$]", conn);
                   
                    objDA.Fill(excelDataSet);
                }
                return Ok(excelDataSet.Tables[0].Rows.Count);
            }
            catch (Exception ex)
            {
                return Ok(ex.Message + " " + ex.StackTrace.ToString());
            }
        }
    }
}
