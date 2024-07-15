using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WebAppExcelReader.Models
{
    public class ExcelData
    {
        public List<string> Columns { get; set; }
        public List<List<string>> Rows { get; set; }
    }
}