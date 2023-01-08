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
using System.Linq;
using LittleBlocks.Excel.ClosedXml.IntegrationTests.Helpers;
using LittleBlocks.Excel.ClosedXml.IntegrationTests.Models;
using LittleBlocks.Excel.Mapper;
using LittleBlocks.Excel.Mapper.PropertyMap;
using LittleBlocks.Excel.Mapper.PropertyMap.Conventions;
using LittleBlocks.Excel.Mapper.PropertyMap.Exceptions;
using LittleBlocks.Excel.Reflection;
using Microsoft.Extensions.Logging;
using NSubstitute;
using Xunit;

namespace LittleBlocks.Excel.ClosedXml.IntegrationTests
{
    public class ExcelPropertyMapFactoryTests : IClassFixture<DatafileFixture>, IDisposable
    {
        public ExcelPropertyMapFactoryTests(DatafileFixture fixture)
        {
            _fixture = fixture;
            _sut = new ExcelPropertyMapFactory(new TypeInfoProvider(), new PropertyInfoExtractor(),
                new PropertyNamingConventionsBucket(new List<IPropertyNameConvention>(),
                    Substitute.For<ILogger<PropertyNamingConventionsBucket>>()),
                Substitute.For<ILogger<ExcelPropertyMapFactory>>());
        }

        public void Dispose()
        {
            _sut = null;
        }

        [Theory]
        [InlineData("FirstName", "FirstName")]
        [InlineData("LastName", "LastName")]
        [InlineData("Age", "Age")]
        public void ShouldCreatePropertyMapFor(string columnName, string propertyName)
        {
            // Arrange
            const string sheetName = "FullDataSheet";
            var worksheet = _fixture.GetWorkbook(Workbook).Worksheet(sheetName);

            // Act
            var result = _sut.CreatePropertyMap<Person>(worksheet);

            // Assert
            var map = result[propertyName];
            Assert.Equal(propertyName, map.Property.Name);
            Assert.Equal(columnName, map.CellValue);
        }

        private void ValidatePropertyMapFor(IDictionary<string, CellToPropertyMap> result, string propertyName,
            string expectedValue)
        {
            var map = result[propertyName];
            Assert.Equal(propertyName, map.Property.Name);
            Assert.Equal(expectedValue, map.CellValue);
        }

        private const string Workbook = "SampleData.xlsx";
        private readonly DatafileFixture _fixture;
        private ExcelPropertyMapFactory _sut;

        [Fact]
        public void OnCreatePropertyMapShouldThrowIfNoWorksheetSupplied()
        {
            // Arrange
            // Act
            // Assert
            Assert.Throws<ArgumentNullException>(() => _sut.CreatePropertyMap<Person>(null));
        }

        [Fact]
        public void ShouldCreatePropertyMapForEntireModel()
        {
            // Arrange
            const string sheetName = "InvalidColumnNameDataSheet";
            var worksheet = _fixture.GetWorkbook(Workbook).Worksheet(sheetName);

            // Act
            var result = _sut.CreatePropertyMap<Person>(worksheet,
                opt => opt.ForMember(p => p.FirstName, Resolve.ByValue("First Name"))
                    .ForMember(p => p.LastName, Resolve.ByValue("Last Name")));

            // Assert
            ValidatePropertyMapFor(result, "FirstName", "First Name");
            ValidatePropertyMapFor(result, "LastName", "Last Name");
            ValidatePropertyMapFor(result, "Age", "Age");
        }

        [Fact]
        public void ShouldIgnoreAgeProperty()
        {
            // Arrange
            const string sheetName = "FullDataSheet";
            var worksheet = _fixture.GetWorkbook(Workbook).Worksheet(sheetName);

            // Act
            var result = _sut.CreatePropertyMap<Person>(worksheet,
                opt => opt.ForMember(p => p.Age, Resolve.Ignore()));

            // Assert
            Assert.Equal(2, result.Count);
            Assert.False(result.ContainsKey("Age"));
        }

        [Fact]
        public void ShouldThrowAggegateExceptionWithTwoInnerExceptionsToCaptureAllMappingProblems()
        {
            // Arrange
            const string sheetName = "FullDataSheet";
            var worksheet = _fixture.GetWorkbook(Workbook).Worksheet(sheetName);

            // Act
            var exception =
                Assert.ThrowsAny<MappingException>(
                    () =>
                    {
                        var result = _sut.CreatePropertyMap<Employee>(worksheet,
                            opt => opt.ForMember(e => e.FirstName, Resolve.ByValue("InvalidPropertyName")));
                    });

            // Assert
            Assert.Equal(2, exception.Exceptions.Count);
            Assert.True(exception.Exceptions.OfType<UnmapppedPropertyException>().Any());
            Assert.True(exception.Exceptions.OfType<MappedColumnByValueNotFoundException>().Any());
        }

        [Fact]
        public void ShouldThrowExceptionForUnmappedProperty()
        {
            // Arrange
            const string sheetName = "FullDataSheet";
            var worksheet = _fixture.GetWorkbook(Workbook).Worksheet(sheetName);

            // Act
            var exception =
                Assert.ThrowsAny<MappingException>(() =>
                {
                    var result = _sut.CreatePropertyMap<Employee>(worksheet);
                });

            // Assert
            Assert.IsType<UnmapppedPropertyException>(exception.Exceptions.Single());
        }

        [Fact]
        public void ShouldThrowIfColumnCantBeResolved()
        {
            // Arrange
            const string sheetName = "FullDataSheet";
            var worksheet = _fixture.GetWorkbook(Workbook).Worksheet(sheetName);

            // Act
            var exception = Assert.ThrowsAny<MappingException>(() =>
            {
                var result = _sut.CreatePropertyMap<Person>(worksheet,
                    opt => opt.ForMember(p => p.FirstName, Resolve.ByValue("InvalidColumn")));
            });

            // Assert
            Assert.IsType<MappedColumnByValueNotFoundException>(exception.Exceptions.Single());
        }

        [Fact]
        public void ShouldThrowIfPropertiesExtractorIsNotSupplied()
        {
            // Arrange
            // Act
            Action action =
                () =>
                    new ExcelPropertyMapFactory(null, new PropertyInfoExtractor(),
                        new PropertyNamingConventionsBucket(new List<IPropertyNameConvention>(),
                            Substitute.For<ILogger<PropertyNamingConventionsBucket>>()),
                        Substitute.For<ILogger<ExcelPropertyMapFactory>>());

            // Assert
            Assert.Throws<ArgumentNullException>(action);
        }

        [Fact]
        public void ShouldThrowIfPropertyInfoExtractorIsNotSupplied()
        {
            // Arrange
            // Act
            Action action =
                () =>
                    new ExcelPropertyMapFactory(new TypeInfoProvider(), null,
                        new PropertyNamingConventionsBucket(new List<IPropertyNameConvention>(),
                            Substitute.For<ILogger<PropertyNamingConventionsBucket>>()),
                        Substitute.For<ILogger<ExcelPropertyMapFactory>>());

            // Assert
            Assert.Throws<ArgumentNullException>(action);
        }

        [Fact]
        public void ShouldThrowIfPropertyNamingConvetionsApplierIsNotSupplied()
        {
            // Arrange
            // Act
            Action action =
                () => new ExcelPropertyMapFactory(new TypeInfoProvider(), new PropertyInfoExtractor(), null,
                    Substitute.For<ILogger<ExcelPropertyMapFactory>>());

            // Assert
            Assert.Throws<ArgumentNullException>(action);
        }
    }
}