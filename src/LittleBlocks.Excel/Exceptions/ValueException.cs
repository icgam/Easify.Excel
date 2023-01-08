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

namespace LittleBlocks.Excel.Exceptions
{
    public sealed class ValueException : ExcelException
    {
        private ValueException(string message)
            : base(message)
        {
        }

        private ValueException(string message, Exception exception)
            : base(message, exception)
        {
        }

        public static ValueException ForMissingValue<TDestinationValueType>(string worksheetName, IDataSheetCell cell)
        {
            var message = $"Value of type '{typeof(TDestinationValueType).Name}' expected in worksheet {worksheetName} cell {cell.ColumnLetter}{cell.RowNumber}";
            return new ValueException(message);
        }

        public static ValueException ForInvalidValue<TDestinationValueType>(IDataSheetCell cell)
        {
            return ForInvalidValue(typeof(TDestinationValueType), cell);
        }

        public static ValueException ForInvalidValue<TDestinationValueType>(IDataSheetCell cell, Exception exception)
        {
            return ForInvalidValue(typeof(TDestinationValueType), cell, exception);
        }

        public static ValueException ForInvalidValue(Type targetType, IDataSheetCell cell)
        {
            if (targetType == null) throw new ArgumentNullException(nameof(targetType));
            if (cell == null) throw new ArgumentNullException(nameof(cell));

            var message = GetMessage(targetType, cell);
            return new ValueException(message);
        }

        public static ValueException ForInvalidValue(Type targetType, IDataSheetCell cell, Exception exception)
        {
            if (targetType == null) throw new ArgumentNullException(nameof(targetType));
            if (cell == null) throw new ArgumentNullException(nameof(cell));
            if (exception == null) throw new ArgumentNullException(nameof(exception));

            var message = GetMessage(targetType, cell);
            return new ValueException(message, exception);
        }

        private static string GetMessage(Type targetType, IDataSheetCell cell)
        {
            return
                $"Conversion from '{typeof(string).Name}' to '{targetType.Name}' for value '{cell.GetString()}' has failed!" +
                $" Value in the cell '{cell.ColumnLetter + cell.RowNumber}' is not valid for the destination type! Please check converter logic or amend the value within the cell.";
        }
    }
}