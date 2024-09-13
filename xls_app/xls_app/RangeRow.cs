using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xls_app
{
    /// <summary>
    /// Диапазон строк в таблице данных
    /// </summary>
    internal class RangeRow
    {
        public uint FirstRow {  get; set; }
        public uint RowsAmount { get; set; }
        public State state { get; set; }

        public enum State : byte
        {
            Single,
            Several,
            All
        }

        public RangeRow(uint firstRow)
        {
            FirstRow = firstRow;
            RowsAmount = 0;
            state = State.Single;
        }

        public RangeRow(uint firstRow, uint rowsNumber)
        {
            FirstRow = firstRow;
            RowsAmount = rowsNumber;
            state = State.Several;
        }

        public RangeRow()
        {
            FirstRow = 0;
            RowsAmount = 0;
            state = State.All;
        }
    }
}
