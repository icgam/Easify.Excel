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
using LittleBlocks.Excel.Exceptions;
using LittleBlocks.Excel.Mapper;
using LittleBlocks.Excel.Mapper.Implementation;
using NSubstitute;
using Xunit;

namespace LittleBlocks.Excel.UnitTests
{
    public class ExcelValueConverterTests : IDisposable
    {
        public ExcelValueConverterTests()
        {
            _converterProvider = Substitute.For<IConverterProvider>();
            _sut = new CellValueConverter(_converterProvider);
        }

        public void Dispose()
        {
            _sut = null;
            _converterProvider = null;
        }

        private IConverterProvider _converterProvider;
        private ICellValueConverter _sut;

        [Fact]
        public void ShouldThrowIConvertersProviderIsNotSupplied()
        {
            // Arrange
            // Act
            Action action = () => new CellValueConverter(null);

            // Assert
            Assert.Throws<ArgumentNullException>(action);
        }

        [Fact]
        public void ShouldThrowIfCellIsNotSupplied()
        {
            // Arrange
            // Act
            Action action = () => _sut.ConvertTo<string>(null);

            // Assert
            Assert.Throws<ArgumentNullException>(action);
        }

        [Fact]
        public void ShouldThrowIfConversionFailsBecauseOfAnInvalidValue()
        {
            // Arrange
            var cell = Substitute.For<IDataSheetCell>();
            cell.GetString().Returns("invalidNumber:)");
            cell.ColumnLetter.Returns("A");
            cell.RowNumber.Returns(1);

            var converters = new Dictionary<Type, Func<IDataSheetCell, object>>();
            Func<IDataSheetCell, object> converterForInt = cellToConvert => int.Parse(cellToConvert.GetString());
            converters.Add(typeof(int), converterForInt);
            _converterProvider.GetConverters().Returns(converters);

            // Act
            Action action = () => _sut.ConvertTo<int>(cell);

            // Assert
            var exception = Assert.Throws<ValueException>(action);
            Assert.Contains("A1", exception.Message);
        }

        [Fact]
        public void ShouldThrowIfConversionIsNotSupported()
        {
            // Arrange
            var cell = Substitute.For<IDataSheetCell>();
            var converters = new Dictionary<Type, Func<IDataSheetCell, object>>();
            Func<IDataSheetCell, object> converterForInt = cellToConvert => default(int);
            converters.Add(typeof(int), converterForInt);
            _converterProvider.GetConverters().Returns(converters);

            // Act
            Action action = () => _sut.ConvertTo<string>(cell);

            // Assert
            Assert.Throws<CellValueConversionNotSupportedException>(action);
        }

        [Fact]
        public void ShouldThrowIfNoConvertersFound()
        {
            // Arrange
            var cell = Substitute.For<IDataSheetCell>();
            _converterProvider.GetConverters().Returns(new Dictionary<Type, Func<IDataSheetCell, object>>());

            // Act
            Action action = () => _sut.ConvertTo<string>(cell);

            // Assert
            Assert.Throws<NoConvertersConfiguredException>(action);
        }

        [Fact]
        public void ShouldThrowIfTargetTypeNotSupplied()
        {
            // Arrange
            var cell = Substitute.For<IDataSheetCell>();

            // Act
            Action action = () => _sut.ConvertTo(cell, null);

            // Assert
            Assert.Throws<ArgumentNullException>(action);
        }
    }
}