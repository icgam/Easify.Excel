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

using LittleBlocks.Excel.Mapper.Implementation;
using LittleBlocks.Testing;
using NSubstitute;
using Xunit;

namespace LittleBlocks.Excel.UnitTests
{
    public class HardcodedConvertersProviderTests
    {
        [Theory]
        [AutoSubstituteAndData]
        public void ShouldNotTrimValueWhenConvertingToString(IDataSheetCell cell, HardcodedConverterProvider sut)
        {
            // Arrange
            cell.GetStringOrDefault().Returns(" Value  ");
            var converters = sut.GetConverters();
            var converter = converters[typeof(string)];
            // Act
            var result = converter(cell);
            // Assert
            Assert.Equal(" Value  ", result);
        }
    }
}