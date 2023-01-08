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

namespace LittleBlocks.Excel.UnitTests.Helpers
{
    internal class FakeDataSheetRow : IDataSheetRow
    {
        private readonly IEnumerable<IDataSheetCell> _cells;

        public FakeDataSheetRow(IEnumerable<IDataSheetCell> cells)
        {
            _cells = cells;
        }

        public string GetStringOrDefault(int col)
        {
            throw new NotImplementedException();
        }

        public string GetDisplayedValue(int col)
        {
            throw new NotImplementedException();
        }

        public DateTime? GetDateTimeOrDefault(int col)
        {
            throw new NotImplementedException();
        }

        public int? GetIntegerOrDefault(int col)
        {
            throw new NotImplementedException();
        }

        public int GetIntegerMandatory(int col)
        {
            throw new NotImplementedException();
        }

        public long? GetLongOrDefault(int col)
        {
            throw new NotImplementedException();
        }

        public bool? GetBooleanOrDefault(int col)
        {
            throw new NotImplementedException();
        }

        public decimal? GetDecimalOrDefault(int col)
        {
            throw new NotImplementedException();
        }

        public bool GetBooleanMandatory(int col)
        {
            throw new NotImplementedException();
        }

        public IEnumerable<IDataSheetCell> Cells()
        {
            return _cells;
        }

        public int RowNumber { get; }

        public IDataSheetCell Cell(string columnLetter)
        {
            throw new NotImplementedException();
        }

        public IDataSheetCell Cell(int columnNumber)
        {
            throw new NotImplementedException();
        }

        public IDataSheetRow Hide()
        {
            throw new NotImplementedException();
        }
    }
}