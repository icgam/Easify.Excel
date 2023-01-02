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
using ClosedXML.Excel;
using LittleBlocks.Excel.ClosedXml.Extensions;

namespace LittleBlocks.Excel.ClosedXml
{
    public class DataSheetCell : IDataSheetCell
    {
        private readonly IXLCell _cell;

        public DataSheetCell(IXLCell cell)
        {
            _cell = cell ?? throw new ArgumentNullException(nameof(cell));
        }

        internal IXLCell Internal => _cell;

        public TValue GetValue<TValue>()
        {
            return _cell.GetValue<TValue>();
        }

        public string GetString()
        {
            return _cell.GetString();
        }

        public string GetStringOrDefault()
        {
            return _cell.GetStringOrDefault();
        }

        public int? GetIntegerOrDefault()
        {
            return _cell.GetIntegerOrDefault();
        }

        public decimal? GetDecimalOrDefault()
        {
            return _cell.GetDecimalOrDefault();
        }

        public long? GetLongOrDefault()
        {
            return _cell.GetLongOrDefault();
        }

        public DateTime? GetDateTimeOrDefault()
        {
            return _cell.GetDateTimeOrDefault();
        }

        public bool? GetBooleanOrDefault()
        {
            return _cell.GetBooleanOrDefault();
        }

        public int GetIntegerMandatory()
        {
            return _cell.GetIntegerMandatory();
        }

        public decimal GetDecimalMandatory()
        {
            return _cell.GetDecimalMandatory();
        }

        public long GetLongMandatory()
        {
            return _cell.GetLongMandatory();
        }

        public DateTime GetDateTimeMandatory()
        {
            return _cell.GetDateTimeMandatory();
        }

        public bool GetBooleanMandatory()
        {
            return _cell.GetBooleanMandatory();
        }

        public Guid GetGuidMandatory()
        {
            return _cell.GetGuidMandatory();
        }

        public Guid? GetGuidOrDefault()
        {
            return _cell.GetGuidOrDefault();
        }

        public object Value
        {
            get => _cell.Value;
            set => _cell.Value = value;
        }

        public bool IsEmpty()
        {
            return _cell.IsEmpty();
        }

        public IDataSheetRow Row()
        {
            return new DataSheetRow(_cell.WorksheetRow());
        }

        public string ColumnLetter => _cell.WorksheetColumn().ColumnLetter();

        public int RowNumber => _cell.WorksheetRow().RowNumber();

        public override string ToString()
        {
            return GetValue<string>();
        }
    }
}