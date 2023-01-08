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
using System.Collections.Generic;

namespace LittleBlocks.Excel.Mapper.Implementation
{
    public sealed class HardcodedConverterProvider : IConverterProvider
    {
        public IReadOnlyDictionary<Type, Func<IDataSheetCell, object>> GetConverters()
        {
            var convertersMap = new Dictionary<Type, Func<IDataSheetCell, object>>
            {
                {typeof(string), cell => cell.GetStringOrDefault()},
                {typeof(int?), cell => cell.GetIntegerOrDefault()},
                {typeof(decimal?), cell => cell.GetDecimalOrDefault()},
                {typeof(long?), cell => cell.GetLongOrDefault()},
                {typeof(DateTime?), cell => cell.GetDateTimeOrDefault()},
                {typeof(bool?), cell => cell.GetBooleanOrDefault()},
                {typeof(int), cell => cell.GetIntegerMandatory()},
                {typeof(decimal), cell => cell.GetDecimalMandatory()},
                {typeof(long), cell => cell.GetLongMandatory()},
                {typeof(DateTime), cell => cell.GetDateTimeMandatory()},
                {typeof(bool), cell => cell.GetBooleanMandatory()},
                {typeof(Guid), cell => cell.GetGuidMandatory()},
                {typeof(Guid?), cell => cell.GetGuidOrDefault()}
            };

            return convertersMap;
        }
    }
}