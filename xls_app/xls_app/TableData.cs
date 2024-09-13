using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xls_app
{
    internal class TableData
    {
        public List<TableDataInstance> GetTableData(string path, string tableName)
        {
            List<TableDataInstance> result = new List<TableDataInstance>();

            DataSheet dataSheet = new DataSheet();
            var sheet = dataSheet.GetDataSheet(path, tableName);
            var rows = sheet.DataRange.RowsUsed();
            var columns = sheet.ColumnsUsed().Skip(1);

            foreach (var row in rows)
            {
                foreach (var column in columns)
                {
                    TableDataInstance instance = new TableDataInstance();
                    instance.NewFileName = row.Cell(1).Value.ToString();
                    instance.ParameterValue = sheet.Cell(row.RowNumber(), column.ColumnNumber()).Value.ToString();
                    int colIndex = column.ColumnNumber();
                    instance.ParameterName = sheet.Fields.ElementAt(colIndex - 1).HeaderCell.Value.ToString();

                    result.Add(instance);
                }
            }
            return result;
        }

        public List<TableDataInstance> GetTableData(string path, string tableName, uint firstRow)
        {
            return null;
        }

        public List<TableDataInstance> GetTableData(string path, string tableName, uint firstRow, uint rowsAmount)
        {
            return null;
        }


    }
}
