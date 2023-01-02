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
using System.Linq;
using LittleBlocks.Excel.ClosedXml.IntegrationTests.Helpers;
using LittleBlocks.Excel.ClosedXml.IntegrationTests.Models;
using LittleBlocks.Excel.Exceptions;
using LittleBlocks.Excel.Mapper;
using LittleBlocks.Excel.Mapper.Implementation;
using Microsoft.Extensions.Logging;
using NSubstitute;
using Xunit;

namespace LittleBlocks.Excel.ClosedXml.IntegrationTests
{
    public class ExcelMapperDataTests : IClassFixture<DatafileFixture>
    {
        public ExcelMapperDataTests(DatafileFixture fixture)
        {
            var fakeLoggerFactory = Substitute.For<ILoggerFactory>();
            fakeLoggerFactory.CreateLogger(Arg.Any<string>()).Returns(Substitute.For<ILogger>());
            
            var fakeWorkbookLoader = Substitute.For<IWorkbookLoader>();

            var factory = new ExcelMapperBuilder(fakeWorkbookLoader, fakeLoggerFactory);
            _fixture = fixture;
            _sut = factory.Build(_fixture.GetWorkbook(Workbook));
        }

        private readonly DatafileFixture _fixture;
        private IExcelMapper _sut;

        private const string Workbook = "SampleData.xlsx";
        private const string FullDataSheetName = "FullDataSheet";
        private const string DataWithGapsSheetName = "MissingValuesDataSheet";
        private const string FormulaSheetName = "FormulaDataSheet";

        [Fact]
        public void ShouldHaveFullMetadataForFirstPerson()
        {
            // Arrange

            // Act
            var result = _sut.Map<PersonWithMetadata>(FullDataSheetName).ToList();

            // Assert
            var metadata = result[0].GetCellMetadata(p => p.FirstName);

            Assert.Equal("A", metadata.ColumnLetter);
            Assert.Equal(2, metadata.RowNumber);
            Assert.Equal("FullDataSheet", metadata.SheetName);
            Assert.Equal("SampleData.xlsx", metadata.FileName);
        }

        [Fact]
        public void ShouldReadFormulasFromDataSheet()
        {
            var worksheet = _fixture.GetWorkbook(Workbook).Worksheet(FormulaSheetName);
            var numberRange = worksheet.Range("A2:F2");

            Assert.Equal(1800, numberRange.Cell(1, 1).GetDecimalMandatory());
            Assert.Equal(1700, numberRange.Cell(1, 2).GetDecimalOrDefault());
            Assert.Equal(1600, numberRange.Cell(1, 3).GetDecimalMandatory());
            Assert.Equal(1500, numberRange.Cell(1, 4).GetDecimalOrDefault());
            Assert.Equal(1400, numberRange.Cell(1, 5).GetDecimalMandatory());
            Assert.Equal(1300, numberRange.Cell(1, 6).GetDecimalOrDefault());
            Assert.Null(numberRange.Cell(1, 7).GetDecimalOrDefault());
            Assert.Throws<ValueException>(() => { numberRange.Cell(1, 7).GetDecimalMandatory(); });

            Assert.Equal(1800, numberRange.Cell(1, 1).GetIntegerMandatory());
            Assert.Equal(1700, numberRange.Cell(1, 2).GetIntegerOrDefault());
            Assert.Equal(1600, numberRange.Cell(1, 3).GetIntegerMandatory());
            Assert.Equal(1500, numberRange.Cell(1, 4).GetIntegerOrDefault());
            Assert.Equal(1400, numberRange.Cell(1, 5).GetIntegerMandatory());
            Assert.Equal(1300, numberRange.Cell(1, 6).GetIntegerOrDefault());
            Assert.Null(numberRange.Cell(1, 7).GetIntegerOrDefault());
            Assert.Throws<ValueException>(() => { numberRange.Cell(1, 7).GetIntegerMandatory(); });

            Assert.Equal(1800, numberRange.Cell(1, 1).GetLongMandatory());
            Assert.Equal(1700, numberRange.Cell(1, 2).GetLongOrDefault());
            Assert.Equal(1600, numberRange.Cell(1, 3).GetLongMandatory());
            Assert.Equal(1500, numberRange.Cell(1, 4).GetLongOrDefault());
            Assert.Equal(1400, numberRange.Cell(1, 5).GetLongMandatory());
            Assert.Equal(1300, numberRange.Cell(1, 6).GetLongOrDefault());
            Assert.Null(numberRange.Cell(1, 7).GetLongOrDefault());
            Assert.Throws<ValueException>(() => { numberRange.Cell(1, 7).GetLongMandatory(); });

            var booleanRange = worksheet.Range("A4:F4");
            Assert.True(booleanRange.Cell(1, 1).GetBooleanMandatory());
            Assert.True(booleanRange.Cell(1, 2).GetBooleanOrDefault());
            Assert.False(booleanRange.Cell(1, 3).GetBooleanMandatory());
            Assert.False(booleanRange.Cell(1, 4).GetBooleanOrDefault());
            Assert.True(booleanRange.Cell(1, 5).GetBooleanMandatory());
            Assert.True(booleanRange.Cell(1, 6).GetBooleanOrDefault());
            Assert.Null(booleanRange.Cell(1, 7).GetBooleanOrDefault());
            Assert.Throws<ValueException>(() => { booleanRange.Cell(1, 7).GetBooleanMandatory(); });

            var stringRange = worksheet.Range("A6:F6");
            Assert.Equal("2000", stringRange.Cell(1, 1).GetStringOrDefault());
            Assert.Equal("1900", stringRange.Cell(1, 2).GetStringOrDefault());
            Assert.Equal("1800", stringRange.Cell(1, 3).GetStringOrDefault());
            Assert.Equal("1700", stringRange.Cell(1, 4).GetStringOrDefault());
            Assert.Equal("1600", stringRange.Cell(1, 5).GetStringOrDefault());
            Assert.Equal("1500", stringRange.Cell(1, 6).GetStringOrDefault());

            var dateRange = worksheet.Range("A8:F8");
            Assert.Equal(new DateTime(2016, 1, 1), dateRange.Cell(1, 1).GetDateTimeMandatory());
            Assert.Equal(new DateTime(2016, 1, 2), dateRange.Cell(1, 2).GetDateTimeOrDefault());
            Assert.Equal(new DateTime(2016, 1, 3), dateRange.Cell(1, 3).GetDateTimeMandatory());
            Assert.Equal(new DateTime(2016, 1, 4), dateRange.Cell(1, 4).GetDateTimeOrDefault());
            Assert.Equal(new DateTime(2016, 1, 5), dateRange.Cell(1, 5).GetDateTimeMandatory());
            Assert.Equal(new DateTime(2016, 1, 6), dateRange.Cell(1, 6).GetDateTimeOrDefault());
            Assert.Null(dateRange.Cell(1, 7).GetDateTimeOrDefault());
            Assert.Throws<ValueException>(() => { dateRange.Cell(1, 7).GetDateTimeMandatory(); });
        }

        [Fact]
        public void ShouldReadFourRecordsWithMissingValuesFromDataSheet()
        {
            // Arrange

            // Act
            var result = _sut.Map<Person>(DataWithGapsSheetName).ToList();

            // Assert
            Assert.Equal(4, result.Count);

            Assert.Equal("John", result[0].FirstName);
            Assert.Equal("Doe", result[0].LastName);
            Assert.Equal(30, result[0].Age);

            Assert.Equal("Jane", result[1].FirstName);
            Assert.Equal("Doe", result[1].LastName);
            Assert.Equal(27, result[1].Age);

            Assert.Equal(string.Empty, result[2].FirstName);
            Assert.Equal("Doe", result[2].LastName);
            Assert.Null(result[2].Age);

            Assert.Equal("Mike", result[3].FirstName);
            Assert.Equal(string.Empty, result[3].LastName);
            Assert.Equal(50, result[3].Age);
        }

        [Fact]
        public void ShouldReadTwoRecordsFromDataSheet()
        {
            // Arrange

            // Act
            var result = _sut.Map<Person>(FullDataSheetName).ToList();

            // Assert
            Assert.Equal(2, result.Count);

            Assert.Equal("John", result[0].FirstName);
            Assert.Equal("Doe", result[0].LastName);
            Assert.Equal(30, result[0].Age);

            Assert.Equal("Jane", result[1].FirstName);
            Assert.Equal("Doe", result[1].LastName);
            Assert.Equal(27, result[1].Age);
        }
    }
}