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
using LittleBlocks.Excel.Mapper.Implementation;
using LittleBlocks.Excel.Reflection;
using LittleBlocks.Excel.UnitTests.Models;
using NSubstitute;
using Xunit;

namespace LittleBlocks.Excel.UnitTests
{
    public class ModelFactoryTests
    {
        [Fact]
        public void ShouldCreatePersonModel()
        {
            // Arrange
            var propertyInfoExtractor = Substitute.For<IPropertyInfoExtractor>();
            var sut = new ModelBuilder(propertyInfoExtractor);

            // Act
            var result = sut.Build<Person>();

            // Assert
            Assert.NotNull(result);
            Assert.IsType<Person>(result);
        }

        [Fact]
        public void ShouldCreatePersonModelWithMetadata()
        {
            // Arrange
            var propertyInfoExtractor = Substitute.For<IPropertyInfoExtractor>();
            var sut = new ModelBuilder(propertyInfoExtractor);

            // Act
            var result = sut.Build<PersonWithMetadata>();

            // Assert
            Assert.NotNull(result);
            Assert.IsType<PersonWithMetadata>(result);
        }

        [Fact]
        public void ShouldThrowIfPropertyInfoExtractorIsNotSupplied()
        {
            // Arrange
            // Act
            Action action = () => new ModelBuilder(null);

            // Assert
            Assert.Throws<ArgumentNullException>(action);
        }
    }
}