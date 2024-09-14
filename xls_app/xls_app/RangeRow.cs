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
        public uint LastRow { get; set; }
        public State rangeState { get; set; }

        public enum State : byte
        {
            Single,
            Several,
            All
        }

        public RangeRow(uint firstRow)
        {
            FirstRow = firstRow;
            LastRow = firstRow + 1;
            rangeState = State.Single;
        }

        public RangeRow(uint firstRow, uint lastRow)
        {
            FirstRow = firstRow;
            LastRow = lastRow;
            rangeState = State.Several;
        }

        public RangeRow()
        {
            rangeState = State.All;
        }
    }
}
