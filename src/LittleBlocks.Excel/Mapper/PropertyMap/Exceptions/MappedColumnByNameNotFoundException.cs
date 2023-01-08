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

using System.Linq;

namespace LittleBlocks.Excel.Mapper.PropertyMap.Exceptions
{
    public sealed class MappedColumnByNameNotFoundException : MappingException
    {
        public MappedColumnByNameNotFoundException(string columnName, IDataSheetRow headerRow)
            : base(GetMesssage(columnName, headerRow))
        {
        }

        private static string GetMesssage(string columnName, IDataSheetRow row)
        {
            var columnNames = row.Cells().Select(c => c.ColumnLetter).ToList();
            return
                $"Cell with column name of '{columnName}' was not found within '{string.Join(", ", columnNames)}'!";
        }
    }
}