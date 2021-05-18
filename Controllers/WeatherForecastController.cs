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
                if (Path.GetExtension(FilePath).Equals(".xls"))
                    reader = ExcelReaderFactory.CreateBinaryReader(filestream);
                else if (Path.GetExtension(FilePath).Equals(".xlsx"))
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
                    return Ok(workbook.MaxColumnCount);
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

                using (OleDbConnection conn = new OleDbConnection(connString))
                {
                    conn.Open();
                    OleDbDataAdapter objDA = new System.Data.OleDb.OleDbDataAdapter
                    ("select * from [Sheet1$]", conn);
                    DataSet excelDataSet = new DataSet();
                    objDA.Fill(excelDataSet);
                }
                return Ok();
            }
            catch (Exception ex)
            {
                return Ok(ex.Message + " " + ex.StackTrace.ToString());
            }
        }
    }
}
