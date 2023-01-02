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
using LittleBlocks.Excel.ClosedXml.Extensions;

namespace LittleBlocks.Excel.ClosedXml
{
    public class DataSheetRow : IDataSheetRow
    {
        private readonly Func<string, IXLCell> _getCellByColumnLetter;
        private readonly Func<int, IXLCell> _getCellByColumnNumber;
        private readonly Func<int> _getRowNumber;
        private readonly Func<IXLRow> _hideRow;
        private readonly IXLRangeBase _row;


        public DataSheetRow(IXLRangeRow row)
        {
            _row = row ?? throw new ArgumentNullException(nameof(row));
            _getCellByColumnNumber = row.Cell;
            _getCellByColumnLetter = row.Cell;
            _getRowNumber = row.RowNumber;
            _hideRow = row.WorksheetRow().Hide;
        }

        public DataSheetRow(IXLRow row)
        {
            _row = row ?? throw new ArgumentNullException(nameof(row));
            _getCellByColumnNumber = row.Cell;
            _getCellByColumnLetter = row.Cell;
            _getRowNumber = row.RowNumber;
            _hideRow = row.Hide;
        }

        public string GetStringOrDefault(int col)
        {
            return _getCellByColumnNumber(col).GetStringOrDefault();
        }

        public string GetDisplayedValue(int col)
        {
            return _getCellByColumnNumber(col).GetDisplayedValue();
        }

        public DateTime? GetDateTimeOrDefault(int col)
        {
            return _getCellByColumnNumber(col).GetDateTimeOrDefault();
        }

        public int? GetIntegerOrDefault(int col)
        {
            return _getCellByColumnNumber(col).GetIntegerOrDefault();
        }

        public int GetIntegerMandatory(int col)
        {
            return _getCellByColumnNumber(col).GetIntegerMandatory();
        }

        public long? GetLongOrDefault(int col)
        {
            return _getCellByColumnNumber(col).GetLongOrDefault();
        }

        public bool? GetBooleanOrDefault(int col)
        {
            return _getCellByColumnNumber(col).GetBooleanOrDefault();
        }

        public decimal? GetDecimalOrDefault(int col)
        {
            return _getCellByColumnNumber(col).GetDecimalOrDefault();
        }

        public bool GetBooleanMandatory(int col)
        {
            return _getCellByColumnNumber(col).GetBooleanMandatory();
        }

        public IEnumerable<IDataSheetCell> Cells()
        {
            // TODO: Cache the cells
            return _row.Cells().Select(c => new DataSheetCell(c));
        }

        public int RowNumber => _getRowNumber();

        public IDataSheetCell Cell(string columnLetter)
        {
            return new DataSheetCell(_getCellByColumnLetter(columnLetter));
        }

        public IDataSheetCell Cell(int columnNumber)
        {
            return new DataSheetCell(_getCellByColumnNumber(columnNumber));
        }

        public IDataSheetRow Hide()
        {
            return new DataSheetRow(_hideRow());
        }
    }
}