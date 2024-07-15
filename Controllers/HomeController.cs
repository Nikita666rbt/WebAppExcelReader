using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using WebAppExcelReader.Models;

namespace WebAppExcelReader.Controllers
{
    public class HomeController : Controller
    {
        // GET: Home
        public ActionResult Index()
        {
            string filePath = Server.MapPath("~/App_Data/Пример.xlsx");
            ExcelData data = ReadExcel(filePath);
            return View(data);
        }

        public ExcelData ReadExcel(string filePath)
        {
            ExcelData data = new ExcelData() {Columns = new List<string>(), Rows = new List<List<string>>()};
            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int colCount = worksheet.Dimension.Columns;
                int rowCount = worksheet.Dimension.Rows;

                for (int col = 1; col <= colCount; col++)
                {
                    data.Columns.Add(worksheet.Cells[1, col].Text);
                }

                for (int row = 2; row <= rowCount; row++)
                {
                    List<string> rowData = new List<string>();
                    for (int col = 1; col <= colCount; col++)
                    {
                        rowData.Add(worksheet.Cells[row, col].Text);
                    }
                    data.Rows.Add(rowData);
                }
            }
            return data;
        }
    }
}