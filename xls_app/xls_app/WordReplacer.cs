using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Words.NET;

namespace xls_app
{
    internal class WordReplacer
    {
        public void WordReplaceValue(string sourceValue, string destinationValue, string filePath)
        {
            using (var document = DocX.Load(filePath))
            {
                document.ReplaceText(sourceValue, destinationValue);
                document.Save();
            }
        }
    }
}
