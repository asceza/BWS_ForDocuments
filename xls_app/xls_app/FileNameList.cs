using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xls_app
{
    internal class FileNameList
    {
        public List<string> GetFileNameList(List<TableDataInstance> tableData )
        {
            
            var fileNames = tableData.Select(x => x.NewFileName)
                .Distinct()
                .ToList();

            return fileNames;
        }
    }
}
