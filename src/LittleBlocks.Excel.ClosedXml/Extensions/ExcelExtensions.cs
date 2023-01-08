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
using System.Globalization;
using ClosedXML.Excel;
using LittleBlocks.Excel.Exceptions;
using LittleBlocks.Extensions;

namespace LittleBlocks.Excel.ClosedXml.Extensions
{
    public static class ExcelExtensions
    {
        public static string GetStringOrDefault(this IXLCell cell)
        {
            try
            {
                return cell.HasFormula ? cell.CachedValue.ToString() : cell.GetString();
            }
            catch (Exception)
            {
                return cell.CachedValue.ToString();
            }
        }

        public static string GetStringMandatory(this IXLCell cell)
        {
            var value = cell.GetStringOrDefault();
            if (value.Empty() || string.IsNullOrEmpty(value) || string.IsNullOrWhiteSpace(value))
                ThrowValueExpectedException<string>(cell);

            return value;
        }

        public static string GetDisplayedValue(this IXLCell cell)
        {
            var decimalValue = GetDecimalOrDefault(cell);
            if (!decimalValue.HasValue || decimalValue.Value == 0 && string.IsNullOrWhiteSpace(cell.Value.ToString()))
                return cell.GetStringOrDefault();

            if (cell.Style.NumberFormat != null &&
                !string.IsNullOrWhiteSpace(cell.Style.NumberFormat.Format))
                return decimalValue.Value.ToString(cell.Style.NumberFormat.Format);

            return decimalValue.Value.ToString(CultureInfo.InvariantCulture);
        }

        public static DateTime? GetDateTimeOrDefault(this IXLCell cell)
        {
            var rawValue = cell.GetStringOrDefault();
            if (string.IsNullOrEmpty(rawValue))
                return null;

            if (double.TryParse(rawValue, out var date))
                return DateTime.FromOADate(date);

            if (DateTime.TryParse(rawValue, out var value))
                return value;

            throw ValueException.ForInvalidValue<DateTime>(new DataSheetCell(cell));
        }

        public static DateTime GetDateTimeMandatory(this IXLCell cell)
        {
            var value = cell.GetDateTimeOrDefault();
            if (value.Empty() || value.HasValue == false) ThrowValueExpectedException<DateTime>(cell);

            return value.Value;
        }

        public static int? GetIntegerOrDefault(this IXLCell cell)
        {
            var rawValue = cell.GetStringOrDefault();

            if (string.IsNullOrEmpty(rawValue))
                return null;

            if (int.TryParse(rawValue, out var integerValue))
                return integerValue;

            throw ValueException.ForInvalidValue<int>(new DataSheetCell(cell));
        }

        public static int GetIntegerMandatory(this IXLCell cell)
        {
            var value = cell.GetIntegerOrDefault();
            if (value.Empty() || value.HasValue == false) ThrowValueExpectedException<int>(cell);

            return value.Value;
        }

        public static long? GetLongOrDefault(this IXLCell cell)
        {
            var rawValue = cell.GetStringOrDefault();

            if (string.IsNullOrEmpty(rawValue))
                return null;

            if (long.TryParse(rawValue, out var longValue))
                return longValue;

            throw ValueException.ForInvalidValue<long>(new DataSheetCell(cell));
        }

        public static long GetLongMandatory(this IXLCell cell)
        {
            var value = cell.GetLongOrDefault();
            if (value.Empty() || value.HasValue == false) ThrowValueExpectedException<long>(cell);

            return value.Value;
        }

        public static bool? GetBooleanOrDefault(this IXLCell cell)
        {
            var rawValue = cell.GetStringOrDefault();

            if (string.IsNullOrEmpty(rawValue))
                return null;

            switch (rawValue)
            {
                case "0":
                    return false;
                case "1":
                    return true;
            }

            if (bool.TryParse(rawValue, out var boolValue))
                return boolValue;

            throw ValueException.ForInvalidValue<bool>(new DataSheetCell(cell));
        }


        public static bool GetBooleanMandatory(this IXLCell cell)
        {
            var value = cell.GetBooleanOrDefault();
            if (value.Empty() || value.HasValue == false) ThrowValueExpectedException<bool>(cell);

            return value.Value;
        }

        public static decimal? GetDecimalOrDefault(this IXLCell cell)
        {
            var rawValue = cell.GetStringOrDefault();

            if (string.IsNullOrEmpty(rawValue))
                return null;

            if (decimal.TryParse(rawValue, out var decimalValue))
                return decimalValue;

            if (double.TryParse(rawValue, out var doubleValue))
                return Convert.ToDecimal(doubleValue);

            throw ValueException.ForInvalidValue<decimal>(new DataSheetCell(cell));
        }


        public static decimal GetDecimalMandatory(this IXLCell cell)
        {
            var value = cell.GetDecimalOrDefault();
            if (value.Empty() || value.HasValue == false) ThrowValueExpectedException<decimal>(cell);

            return value.Value;
        }

        public static Guid? GetGuidOrDefault(this IXLCell cell)
        {
            var rawValue = cell.GetStringOrDefault();

            if (string.IsNullOrEmpty(rawValue))
                return null;

            if (Guid.TryParse(rawValue, out var guidValue))
                return guidValue;

            throw ValueException.ForInvalidValue<Guid>(new DataSheetCell(cell));
        }

        public static Guid GetGuidMandatory(this IXLCell cell)
        {
            var value = cell.GetGuidOrDefault();
            if (value.Empty() || value.HasValue == false) ThrowValueExpectedException<Guid>(cell);

            return value.Value;
        }

        private static void ThrowValueExpectedException<TExpectedValueType>(IXLCell cell)
        {
            throw ValueException.ForMissingValue<TExpectedValueType>(cell.Worksheet.Name, new DataSheetCell(cell));
        }
    }
}
