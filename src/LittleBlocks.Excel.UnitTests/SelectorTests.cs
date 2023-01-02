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
using LittleBlocks.Excel.Mapper.PropertyMap.Exceptions;
using LittleBlocks.Excel.Mapper.PropertyMap.Selectors;
using LittleBlocks.Excel.UnitTests.Helpers;
using LittleBlocks.Testing;
using NSubstitute;
using Xunit;

namespace LittleBlocks.Excel.UnitTests
{
    public class SelectorTests
    {
        [Theory]
        [AutoSubstituteAndData]
        public void ShouldThrowIfDuplicateColumnValueFound(List<IDataSheetCell> cells)
        {
            // Arrange
            var sut = new ColumnSelectorByValue("value", StringComparison.CurrentCultureIgnoreCase);
            var row = new FakeDataSheetRow(cells);
            cells[0].GetValue<string>().Returns("value");
            cells[0].ColumnLetter.Returns("A");
            cells[0].RowNumber.Returns(1);
            cells[2].GetValue<string>().Returns("value");
            cells[2].ColumnLetter.Returns("A");
            cells[2].RowNumber.Returns(3);

            // Act
            // Assert
            var exception = Assert.Throws<MappedColumnHasDuplicateValueException>(() => sut.SelectCellOrThrow(row));
            Assert.Equal("Duplicate column header 'value' found at 'A:1, A:3'!", exception.Message);
        }

        [Fact]
        public void ShouldBeTreatedAsIgnoreColumnSelector()
        {
            // Arrange
            var sut = new IgnoreColumn();
            // Act
            // Assert
            Assert.True(sut.Is<IgnoreColumn>());
        }

        [Fact]
        public void ShouldNotBeTreatedAsIgnoreColumnSelector()
        {
            // Arrange
            var sut = new ColumnSelectorByLetter("SomePropertyName");
            // Act
            // Assert
            Assert.False(sut.Is<IgnoreColumn>());
        }
    }
}