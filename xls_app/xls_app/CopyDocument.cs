using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace xls_app
{
    internal class CopyDocument
    {
        public void CopyDoc(List<string> fileNames, string folderPath, string originFilePath)
        {
            foreach (string fileName in fileNames)
            {
                if (originFilePath.Contains(".docx"))
                {
                    string fileType = ".docx";
                    string destinationFilePath = folderPath + fileName + fileType;
                    try
                    {
                        File.Copy(originFilePath, destinationFilePath);
                    }
                    catch (IOException ex)
                    {
                        MessageBox.Show($"Ошибка при копировании файла: {ex.Message}");
                    }
                }
                else if (originFilePath.Contains(".xlsx"))
                {
                    string fileType = ".xlsx";
                    string destinationFilePath = folderPath + fileName + fileType;
                    try
                    {
                        File.Copy(originFilePath, destinationFilePath);
                    }
                    catch (IOException ex)
                    {
                        MessageBox.Show($"Ошибка при копировании файла: {ex.Message}");
                    }
                }
                                

            }
                        
            

        }
    }
}
