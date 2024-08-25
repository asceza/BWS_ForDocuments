using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace xls_app
{
    internal class DataSheet
    {
        public IXLTable GetDataSheet(string path, string tableName)
        {
            IXLWorkbook wb = new XLWorkbook(path);           
            var sheet = wb.Table(tableName);            
            return sheet;
        }
    }
}
