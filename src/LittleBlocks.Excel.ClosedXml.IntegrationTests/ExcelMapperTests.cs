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
using System.Linq.Expressions;
using System.Reflection;
using LittleBlocks.Excel.ClosedXml.IntegrationTests.Helpers;
using LittleBlocks.Excel.ClosedXml.IntegrationTests.Models;
using LittleBlocks.Excel.Exceptions;
using LittleBlocks.Excel.Mapper;
using LittleBlocks.Excel.Mapper.Implementation;
using LittleBlocks.Excel.Mapper.PropertyMap;
using LittleBlocks.Excel.Reflection;
using Microsoft.Extensions.Logging;
using NSubstitute;
using NSubstitute.ReturnsExtensions;
using Xunit;

namespace LittleBlocks.Excel.ClosedXml.IntegrationTests
{
    public class ExcelMapperTests : IClassFixture<DatafileFixture>
    {
        public ExcelMapperTests(DatafileFixture fixture)
        {
            _fixture = fixture;
            _propertyMapFactory = Substitute.For<IExcelPropertyMapFactory>();
            _typeConverter = Substitute.For<ICellValueConverter>();
            _modelBuilder = Substitute.For<IModelBuilder>();
            _propertyInfoExtractor = Substitute.For<IPropertyInfoExtractor>();
            _sut = new ExcelMapper(_fixture.GetWorkbook(Workbook), _propertyMapFactory, _typeConverter, _modelBuilder,
                Substitute.For<ILogger<ExcelMapper>>());
            _propertyMapFactory.CreatePropertyMap<Person>(Arg.Any<IDataSheet>()).Returns(GetMappings());
            _propertyMapFactory.CreatePropertyMap(Arg.Any<IDataSheet>(), Arg.Any<Action<IMemberSpec<Person>>>())
                .Returns(GetMappings());
            _propertyMapFactory.CreatePropertyMap<PersonWithMetadata>(Arg.Any<IDataSheet>())
                .Returns(GetMappingsWithMetadata());
            _propertyMapFactory.CreatePropertyMap(Arg.Any<IDataSheet>(),
                    Arg.Any<Action<IMemberSpec<PersonWithMetadata>>>())
                .Returns(GetMappingsWithMetadata());
            _modelBuilder.Build<Person>().Returns(new Person());
            _modelBuilder.Build<PersonWithMetadata>().Returns(CreateModel(), CreateModel());
        }

        private const string Workbook = "SampleData.xlsx";
        private readonly DatafileFixture _fixture;
        private IExcelPropertyMapFactory _propertyMapFactory;
        private ICellValueConverter _typeConverter;
        private ExcelMapper _sut;
        private IModelBuilder _modelBuilder;
        private IPropertyInfoExtractor _propertyInfoExtractor;

        private const string SheetName = "FullDataSheet";

        private PersonWithMetadata CreateModel()
        {
            return new PersonWithMetadata(_propertyInfoExtractor);
        }

        private IDictionary<string, CellToPropertyMap> GetMappings()
        {
            var mappings = new Dictionary<string, CellToPropertyMap>();
            mappings.Add("FirstName", new CellToPropertyMap(typeof(Person).GetProperty("FirstName"), "A", 1, "YES:)"));
            return mappings;
        }

        private IDictionary<string, CellToPropertyMap> GetMappingsWithMetadata()
        {
            var mappings = new Dictionary<string, CellToPropertyMap>();
            mappings.Add("FirstName",
                new CellToPropertyMap(typeof(PersonWithMetadata).GetProperty("FirstName"), "A", 1, "John"));
            mappings.Add("LastName",
                new CellToPropertyMap(typeof(PersonWithMetadata).GetProperty("LastName"), "B", 1, "Dow"));
            mappings.Add("Age", new CellToPropertyMap(typeof(PersonWithMetadata).GetProperty("Age"), "C", 1, "30"));
            return mappings;
        }

        [Fact]
        public void ShouldBuildColumnMappingsBasedOnPassedinConfiguration()
        {
            // Arrange
            Action<IMemberSpec<Person>> mappings = m => m.ForMember(p => p.FirstName, Resolve.ByValue("MyColumn"));
            // Act
            _sut.Map(SheetName, mappings);

            // Assert
            _propertyMapFactory.Received(1).CreatePropertyMap(Arg.Is<IDataSheet>(w => w.Name == SheetName), mappings);
        }

        [Fact]
        public void ShouldContainMetadataForEachProperty()
        {
            // Arrange
            var propertyInfo = Substitute.For<PropertyInfo>();
            propertyInfo.Name.Returns("FirstName", "LastName", "Age");
            _propertyInfoExtractor.GetPropertyInfo(Arg.Any<Expression<Func<PersonWithMetadata, object>>>())
                .Returns(propertyInfo);

            // Act
            var result = _sut.Map<PersonWithMetadata>(SheetName).First();

            // Assert
            Assert.NotNull(result.GetCellMetadata(p => p.FirstName));
            Assert.NotNull(result.GetCellMetadata(p => p.LastName));
            Assert.NotNull(result.GetCellMetadata(p => p.Age));
        }

        [Fact]
        public void ShouldConvertTypeUsingValueTypeConverter()
        {
            // Arrange
            // Act
            _sut.Map<Person>(SheetName);

            // Assert
            _typeConverter.Received(2)
                .ConvertTo(Arg.Any<IDataSheetCell>(), typeof(Person).GetProperty("FirstName").PropertyType);
        }

        [Fact]
        public void ShouldThrowAggregatedExceptionContainingAllMappingErrors()
        {
            // Arrange
            var cell = Substitute.For<IDataSheetCell>();
            cell.GetString().Returns("value1", "value2");
            cell.ColumnLetter.Returns("A", "B");
            cell.RowNumber.Returns(2, 3);


            _typeConverter.ConvertTo(Arg.Any<IDataSheetCell>(), Arg.Any<Type>())
                .Returns(x => throw ValueException.ForInvalidValue(typeof(string), cell, new Exception()));

            // Act
            Action mapAction = () => _sut.Map<Person>(SheetName);

            // Assert
            var exception = Assert.Throws<WorksheetMappingFailedException>(mapAction);
            Assert.Equal(2, exception.Exceptions.Count);

            var first = exception.Exceptions.First();
            var second = exception.Exceptions.Skip(1).First();

            Assert.True(first.Message.Contains("A2") && first.Message.Contains("value1"));
            Assert.True(second.Message.Contains("B3") && second.Message.Contains("value2"));
        }

        [Fact]
        public void ShouldThrowIfModelFactoryIsNotSupplied()
        {
            // Arrange
            // Act
            Action action = () => new ExcelMapper(_fixture.GetWorkbook(Workbook), _propertyMapFactory, _typeConverter,
                null, Substitute.For<ILogger<ExcelMapper>>());

            // Assert
            Assert.Throws<ArgumentNullException>(action);
        }

        [Fact]
        public void ShouldThrowIfNoMappingConfigurationPassed()
        {
            // Arrange
            // Act
            Action mapAction = () => _sut.Map<Person>(SheetName, (Action<IMemberSpec<Person>>)null);

            // Assert
            Assert.Throws<ArgumentNullException>(mapAction);
        }

        [Fact]
        public void ShouldThrowIfNoMappingsCreated()
        {
            // Arrange
            _propertyMapFactory.CreatePropertyMap(Arg.Any<IDataSheet>(), Arg.Any<Action<IMemberSpec<Person>>>())
                .ReturnsNull();

            // Act
            Action mapAction = () => _sut.Map<Person>(SheetName);

            // Assert
            Assert.Throws<NoMappingsCreatedException<Person>>(mapAction);
        }

        [Fact]
        public void ShouldThrowIfPropertyMapFactoryIsNotSupplied()
        {
            // Arrange
            // Act
            Action action = () => new ExcelMapper(_fixture.GetWorkbook(Workbook), null, _typeConverter, _modelBuilder,
                Substitute.For<ILogger<ExcelMapper>>());

            // Assert
            Assert.Throws<ArgumentNullException>(action);
        }

        [Fact]
        public void ShouldThrowIfTypeConverterIsNotSupplied()
        {
            // Arrange
            // Act
            Action action = () => new ExcelMapper(_fixture.GetWorkbook(Workbook), _propertyMapFactory, null,
                _modelBuilder, Substitute.For<ILogger<ExcelMapper>>());

            // Assert
            Assert.Throws<ArgumentNullException>(action);
        }

        [Fact]
        public void ShouldThrowIfWorkbookIsNotSupplied()
        {
            // Arrange
            // Act
            Action action = () => new ExcelMapper(null, _propertyMapFactory, _typeConverter, _modelBuilder,
                Substitute.For<ILogger<ExcelMapper>>());

            // Assert
            Assert.Throws<ArgumentNullException>(action);
        }

        [Fact]
        public void ShouldThrowIfWorksheetNotFound()
        {
            // Arrange
            const string nonExistingWorksheet = "NonExistingWorksheetName";
            // Act
            Action mapAction = () => _sut.Map<Person>(nonExistingWorksheet);

            // Assert
            Assert.Throws<WorksheetNotFoundException>(mapAction);
        }
    }
}