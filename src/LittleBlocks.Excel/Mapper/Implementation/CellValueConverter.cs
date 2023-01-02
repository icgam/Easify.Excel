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
using LittleBlocks.Excel.Exceptions;
using LittleBlocks.Extensions;

namespace LittleBlocks.Excel.Mapper.Implementation
{
    public sealed class CellValueConverter : ICellValueConverter
    {
        private readonly IConverterProvider _converterProvider;

        public CellValueConverter(IConverterProvider converterProvider)
        {
            _converterProvider = converterProvider ?? throw new ArgumentNullException(nameof(converterProvider));
        }

        public TTargetValue ConvertTo<TTargetValue>(IDataSheetCell cell)
        {
            var value = ConvertTo(cell, typeof(TTargetValue));
            if (value.AnyValue()) return (TTargetValue) value;

            return default(TTargetValue);
        }

        public object ConvertTo(IDataSheetCell cell, Type convertTo)
        {
            if (cell == null) throw new ArgumentNullException(nameof(cell));
            if (convertTo == null) throw new ArgumentNullException(nameof(convertTo));

            var converters = _converterProvider.GetConverters();

            if (converters.Empty()) throw new NoConvertersConfiguredException();

            if (converters.ContainsKey(convertTo) == false)
                throw new CellValueConversionNotSupportedException(convertTo, cell);

            var converter = converters[convertTo];

            try
            {
                return converter(cell);
            }
            catch (Exception exception)
            {
                throw ValueException.ForInvalidValue(convertTo, cell, exception);
            }
        }
    }
}