// This software is part of the LittleBlocks.Excel Library
// Copyright (C) 2018 LittleBlocks
// 
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
// 
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU Affero General Public License for more details.
// 
// You should have received a copy of the GNU Affero General Public License
// along with this program.  If not, see <http://www.gnu.org/licenses/>.
// 

using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

namespace LittleBlocks.Excel.ClosedXml
{
    public class DataSheetRange : IDataSheetRange
    {
        private readonly IXLRange _range;

        public DataSheetRange(IXLRange range)
        {
            _range = range ?? throw new ArgumentNullException(nameof(range));
        }

        public IEnumerable<IDataSheetRow> Rows
        {
            get { return _range.Rows().Select(r => new DataSheetRow(r)); }
        }

        public IDataSheetCell Cell(int row, int column)
        {
            return new DataSheetCell(_range.Cell(row, column));
        }

        public int RowCount => _range.RowCount();

        public IDataSheetRow Row(int row)
        {
            return new DataSheetRow(_range.Row(row));
        }
    }
}