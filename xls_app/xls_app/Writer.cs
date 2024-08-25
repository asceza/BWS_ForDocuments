using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace xls_app
{
    internal class Writer
    {
        public void WriteValue (List<TableDataInstance> valueList, List<string> filesPath, string spChar, string folderPath) 
        {            
            foreach (string file in filesPath)
            {
                if (file.Contains(".docx"))
                {
                    string clearFileName = file.Replace(folderPath, "").Replace(".docx", "");
                    var fileValues = valueList.Where(a => a.NewFileName == clearFileName);

                    foreach (TableDataInstance value in fileValues)
                    {
                        string sourceValue = spChar + value.ParameterName + spChar;
                        string destinationValue = value.ParameterValue;
                        WordReplacer wordReplaser = new WordReplacer();
                        wordReplaser.WordReplaceValue(sourceValue, destinationValue, file);
                    }
                }
                else if (file.Contains(".xlsx"))
                {
                    string clearFileName = file.Replace(folderPath, "").Replace(".xlsx", "");
                    var fileValues = valueList.Where(a => a.NewFileName == clearFileName);

                    foreach (TableDataInstance value in fileValues)
                    {
                        string sourceValue = spChar + value.ParameterName + spChar;
                        string destinationValue = value.ParameterValue;
                        ExcelReplacer er = new ExcelReplacer();
                        er.ExcelReplaceValue(sourceValue, destinationValue, file);                       
                    }

                }
                
            }
                  
        }
    }
}
