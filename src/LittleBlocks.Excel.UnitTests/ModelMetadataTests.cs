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
using LittleBlocks.Excel.Mapper.Metadata.Exceptions;
using LittleBlocks.Excel.Reflection;
using LittleBlocks.Excel.UnitTests.Models;
using Xunit;

namespace LittleBlocks.Excel.UnitTests
{
    public class ModelMetadataTests
    {
        public ModelMetadataTests()
        {
            _sut = new PersonWithMetadata(new PropertyInfoExtractor());
        }

        private readonly PersonWithMetadata _sut;

        [Fact]
        public void ShouldBeAbleToRetrieveStoredMetadataForGivenProperty()
        {
            // Arrange
            var propertyInfo = typeof(PersonWithMetadata).GetProperty(nameof(PersonWithMetadata.FirstName));
            // Act
            _sut.AddCellMetadata(propertyInfo, "Sheet1", "A", 1);

            var meta = _sut.GetCellMetadata(p => p.FirstName);
            // Assert
            Assert.Equal("Sheet1", meta.SheetName);
            Assert.Equal("A", meta.ColumnLetter);
            Assert.Equal(1, meta.RowNumber);
        }

        [Fact]
        public void ShouldIndicateThatCellHasNoMetadataSet()
        {
            // Arrange
            var propertyInfo = typeof(PersonWithMetadata).GetProperty(nameof(PersonWithMetadata.FirstName));
            // Act
            _sut.AddCellMetadata(propertyInfo, "Sheet1", "A", 1);
            var result = _sut.CellHasMetadata(p => p.LastName);

            // Assert
            Assert.False(result);
        }

        [Fact]
        public void ShouldIndicateThatCellMetadataIsSet()
        {
            // Arrange
            var propertyInfo = typeof(PersonWithMetadata).GetProperty(nameof(PersonWithMetadata.FirstName));
            // Act
            _sut.AddCellMetadata(propertyInfo, "Sheet1", "A", 1);
            var result = _sut.CellHasMetadata(p => p.FirstName);

            // Assert
            Assert.True(result);
        }

        [Fact]
        public void ShouldThrowIfMetadataAddedMoreThanOnceForSameProperty()
        {
            // Arrange
            var propertyInfo = typeof(PersonWithMetadata).GetProperty(nameof(PersonWithMetadata.FirstName));
            // Act
            Action action = () =>
            {
                _sut.AddCellMetadata(propertyInfo, "Sheet1", "A", 1);
                _sut.AddCellMetadata(propertyInfo, "Sheet1", "B", 1);
            };

            // Assert
            Assert.Throws<PropertyCanHaveOnlyOneCellMetadataException>(action);
        }

        [Fact]
        public void ShouldThrowIfNoMetadataFoundForProperty()
        {
            // Arrange
            // Act
            Action action = () => { _sut.GetCellMetadata(p => p.FirstName); };

            // Assert
            Assert.Throws<PropertyMetadataNotFoundException>(action);
        }

        [Fact]
        public void ShouldThrowIfPropertyInfoExtractorIsNotSupplied()
        {
            // Arrange
            // Act
            Action action = () => new PersonWithMetadata(null);

            // Assert
            Assert.Throws<ArgumentNullException>(action);
        }
    }
}