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
using System.Linq;
using LittleBlocks.Excel.Mapper.PropertyMap.Exceptions;

namespace LittleBlocks.Excel.Mapper.PropertyMap.Selectors
{
    public class ColumnSelectorByLetter : ColumnSelector
    {
        private readonly string _columnNameToFind;

        public ColumnSelectorByLetter(string columnNameToFind)
        {
            if (string.IsNullOrEmpty(columnNameToFind))
                throw new ArgumentException("Argument is null or empty", nameof(columnNameToFind));
            if (string.IsNullOrWhiteSpace(columnNameToFind))
                throw new ArgumentException("Argument is null or whitespace", nameof(columnNameToFind));

            _columnNameToFind = columnNameToFind.ToUpper();
        }

        public override IDataSheetCell SelectCellOrThrow(IDataSheetRow headerRow)
        {
            var cell = SelectCell(headerRow);
            if (cell == null) throw new MappedColumnByNameNotFoundException(_columnNameToFind, headerRow);

            return cell;
        }

        public override IDataSheetCell SelectCell(IDataSheetRow headerRow)
        {
            if (headerRow == null) throw new ArgumentNullException(nameof(headerRow));

            var cell =
                headerRow.Cells()
                    .SingleOrDefault(
                        c => c.ColumnLetter == _columnNameToFind);

            return cell;
        }
    }
}