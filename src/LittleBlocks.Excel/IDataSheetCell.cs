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

namespace LittleBlocks.Excel
{
    public interface IDataSheetCell : IDataSheetCellCoordinates
    {
        object Value { get; set; }
        TValue GetValue<TValue>();
        object OriginalCell { get; }

        string GetString();
        string GetStringOrDefault();
        int? GetIntegerOrDefault();
        decimal? GetDecimalOrDefault();
        long? GetLongOrDefault();
        DateTime? GetDateTimeOrDefault();
        bool? GetBooleanOrDefault();
        int GetIntegerMandatory();
        decimal GetDecimalMandatory();
        long GetLongMandatory();
        DateTime GetDateTimeMandatory();
        bool GetBooleanMandatory();
        Guid GetGuidMandatory();
        Guid? GetGuidOrDefault();
        bool IsEmpty();

        IDataSheetRow Row();
    }
}