using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;


namespace xls_app
{
    internal class ExcelReplacer
    {
        public void ExcelReplaceValue(string sourceValue, string destinationValue, string filePath)
        {
            using (var workbook = new XLWorkbook(filePath))
            {               
                foreach (var worksheet in workbook.Worksheets)
                {                    
                    foreach (var cell in worksheet.CellsUsed())
                    {                        
                        if (cell.HasFormula == false && cell.GetString().Contains(sourceValue))
                        {                            
                            string newValue = cell.GetString().Replace(sourceValue, destinationValue);
                            cell.Value = newValue;
                        }
                    }
                }                
                workbook.Save();
            }

        }
    }
}
