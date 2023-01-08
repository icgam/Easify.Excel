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
    public class ColumnSelectorByValue : ColumnSelector
    {
        private readonly StringComparison _options;
        private readonly string _valueToFind;

        public ColumnSelectorByValue(string valueToFind) : this(valueToFind, StringComparison.CurrentCulture)
        {
        }

        public ColumnSelectorByValue(string valueToFind, StringComparison options)
        {
            _valueToFind = valueToFind;
            _options = options;
            if (string.IsNullOrEmpty(valueToFind))
                throw new ArgumentException("Argument is null or empty", nameof(valueToFind));
            if (string.IsNullOrWhiteSpace(valueToFind))
                throw new ArgumentException("Argument is null or whitespace", nameof(valueToFind));
            if (!Enum.IsDefined(typeof(StringComparison), options))
                throw new ArgumentOutOfRangeException(nameof(options));
        }

        public override IDataSheetCell SelectCellOrThrow(IDataSheetRow headerRow)
        {
            var cell = SelectCell(headerRow);
            if (cell == null) throw new MappedColumnByValueNotFoundException(_valueToFind, headerRow);

            return cell;
        }

        public override IDataSheetCell SelectCell(IDataSheetRow headerRow)
        {
            if (headerRow == null) throw new ArgumentNullException(nameof(headerRow));

            var matchingCells =
                headerRow.Cells()
                    .Where(
                        c => string.Compare(c.GetValue<string>(), _valueToFind, _options) == 0).ToList();

            if (matchingCells.Count > 1) throw new MappedColumnHasDuplicateValueException(_valueToFind, matchingCells);

            return matchingCells.SingleOrDefault();
        }
    }
}