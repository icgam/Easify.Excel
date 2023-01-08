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

namespace LittleBlocks.Excel.Exceptions
{
    public sealed class NoConvertersConfiguredException : ExcelException
    {
        public NoConvertersConfiguredException()
            : base(GetMessage())
        {
        }

        private static string GetMessage()
        {
            return
                "No type converters configured for conversions from Excel cell value to valid .NET type. Please amend the system configuration and add converters in order to allow loading Excel data files.";
        }
    }
}