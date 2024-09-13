using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xls_app
{
    internal class RangeRow
    {
        public uint FirstRow {  get; set; }
        public uint RowsNumber { get; set; }
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
            RowsNumber = 0;
            state = State.Single;
        }

        public RangeRow(uint firstRow, uint rowsNumber)
        {
            FirstRow = firstRow;
            RowsNumber = rowsNumber;
            state = State.Several;
        }

        public RangeRow()
        {
            FirstRow = 0;
            RowsNumber = 0;
            state = State.All;
        }
    }
}
