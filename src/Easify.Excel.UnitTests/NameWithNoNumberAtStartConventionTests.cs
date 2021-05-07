// This software is part of the Easify.Ef Library
// Copyright (C) 2018 Intermediate Capital Group
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


using Easify.Excel.Mapper.PropertyMap.Convetions;
using Xunit;

namespace Easify.Excel.UnitTests
{
    public class NameWithNoNumberAtStartConventionTests
    {
        [Theory]
        [InlineData("GenCamaro", "GenCamaro")]
        [InlineData("Gen1Camaro", "Gen1Camaro")]
        [InlineData("1Gen1Camaro", "Gen1Camaro")]
        public void ShouldReturnSameValueIsNotStartingWithNumeric(string sourceValue, string expectedValue)
        {
            // Arrange
            var sut = new NameWithNoNumberAtStartConvention();
            
            // Act
            var result = sut.ApplyConvention(sourceValue);

            // Assert
            Assert.Equal(expectedValue, result);
        }

        [Theory]
        [InlineData("0GenCamaro")]
        [InlineData("1GenCamaro")]
        [InlineData("2GenCamaro")]
        [InlineData("3GenCamaro")]
        [InlineData("4GenCamaro")]
        [InlineData("5GenCamaro")]
        [InlineData("6GenCamaro")]
        [InlineData("7GenCamaro")]
        [InlineData("8GenCamaro")]
        [InlineData("9GenCamaro")]
        [InlineData("10GenCamaro")]
        [InlineData("123GenCamaro")]
        public void ShouldReturnValueWithNumericValueTrimmedFromTheStart(string sourceValue)
        {
            // Arrange
            var sut = new NameWithNoNumberAtStartConvention();
            
            // Act
            var result = sut.ApplyConvention(sourceValue);

            // Assert
            Assert.Equal("GenCamaro", result);
        }
    }
}