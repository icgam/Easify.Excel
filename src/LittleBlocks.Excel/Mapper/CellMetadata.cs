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

using LittleBlocks.Extensions;

namespace LittleBlocks.Excel.Mapper
{
    public class CellMetadata : IDataSheetCellCoordinates
    {
        public CellMetadata(string fileName, string sheetName, string columnLetter,
            int rowNumber)
        {
            FileName = fileName;
            SheetName = sheetName;
            ColumnLetter = columnLetter;
            RowNumber = rowNumber;
        }

        public string FileName { get; }
        public string SheetName { get; }
        public string ColumnLetter { get; }
        public int RowNumber { get; }

        public override string ToString()
        {
            return FileName.AnyValue()
                ? $"Cell {ColumnLetter}:{RowNumber} in {FileName} -> {SheetName} datasheet."
                : $"Cell {ColumnLetter}:{RowNumber} in {SheetName} datasheet.";
        }
    }
}