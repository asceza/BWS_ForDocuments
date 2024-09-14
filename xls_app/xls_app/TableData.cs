using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xls_app
{
    internal class TableData
    {
        /// <summary>
        /// Получение данных из таблицы (все строки)
        /// </summary>
        /// <param name="path">Путь к файлу с таблицей</param>
        /// <param name="tableName">Название таблицы в документе</param>
        /// <returns></returns>
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

        /// <summary>
        /// Получение данных из таблицы (одна строка)
        /// </summary>
        /// <param name="path">Путь к файлу с таблицей</param>
        /// <param name="tableName">Название таблицы в документе</param>
        /// <param name="firstRow">Номер строки</param>
        /// <returns></returns>
        public List<TableDataInstance> GetTableData(string path, string tableName, uint firstRow)
        {
            uint currentRow = 0;
            List<TableDataInstance> result = new List<TableDataInstance>();

            DataSheet dataSheet = new DataSheet();
            var sheet = dataSheet.GetDataSheet(path, tableName);
            var rows = sheet.DataRange.RowsUsed();
            var columns = sheet.ColumnsUsed().Skip(1);

            foreach (var row in rows)
            {
                currentRow++;
                foreach (var column in columns)
                {
                    if (currentRow == firstRow)
                    {
                        TableDataInstance instance = new TableDataInstance();
                        instance.NewFileName = row.Cell(1).Value.ToString();
                        instance.ParameterValue = sheet.Cell(row.RowNumber(), column.ColumnNumber()).Value.ToString();
                        int colIndex = column.ColumnNumber();
                        instance.ParameterName = sheet.Fields.ElementAt(colIndex - 1).HeaderCell.Value.ToString();
                        result.Add(instance);
                    }
                    if (currentRow > firstRow)
                    {
                        break;
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// Получение данных из таблицы (несколько строк)
        /// </summary>
        /// <param name="path">Путь к файлу с таблицей</param>
        /// <param name="tableName">Название таблицы в документе</param>
        /// <param name="firstRow">Номер первой строки</param>
        /// <param name="lastRow">Номер последней строки</param>
        /// <returns></returns>
        public List<TableDataInstance> GetTableData(string path, string tableName, uint firstRow, uint lastRow)
        {
            uint currentRow = 0;
            List<TableDataInstance> result = new List<TableDataInstance>();

            DataSheet dataSheet = new DataSheet();
            var sheet = dataSheet.GetDataSheet(path, tableName);
            var rows = sheet.DataRange.RowsUsed();
            var columns = sheet.ColumnsUsed().Skip(1);

            foreach (var row in rows)
            {
                currentRow++;
                foreach (var column in columns)
                {
                    if (currentRow >= firstRow)
                    {
                        TableDataInstance instance = new TableDataInstance();
                        instance.NewFileName = row.Cell(1).Value.ToString();
                        instance.ParameterValue = sheet.Cell(row.RowNumber(), column.ColumnNumber()).Value.ToString();
                        int colIndex = column.ColumnNumber();
                        instance.ParameterName = sheet.Fields.ElementAt(colIndex - 1).HeaderCell.Value.ToString();
                        result.Add(instance);
                    }
                    if (currentRow > lastRow)
                    {
                        break;
                    }
                }
            }
            return result;
        }


    }
}
