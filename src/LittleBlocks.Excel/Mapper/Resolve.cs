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
using LittleBlocks.Excel.Mapper.PropertyMap.Selectors;

namespace LittleBlocks.Excel.Mapper
{
    public static class Resolve
    {
        public static IColumnSelector ByColumnLetter(string columnName)
        {
            return new ColumnSelectorByLetter(columnName);
        }

        public static IColumnSelector ByValue(string value)
        {
            return new ColumnSelectorByValue(value);
        }

        public static IColumnSelector ByValue(string value, StringComparison options)
        {
            return new ColumnSelectorByValue(value, options);
        }

        public static IColumnSelector ByConvention(string value, IPropertyNameConvention convention)
        {
            return new ColumnSelectorByConvention(value, convention);
        }

        public static IColumnSelector ByConvention(string value, IPropertyNameConvention convention,
            StringComparison options)
        {
            return new ColumnSelectorByConvention(value, convention, options);
        }

        public static IColumnSelector Ignore()
        {
            return new IgnoreColumn();
        }
    }
}