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

using System.Diagnostics;
using System.Linq;
using LittleBlocks.Excel.ClosedXml.IntegrationTests.Helpers;
using LittleBlocks.Excel.Mapper;
using LittleBlocks.Excel.Mapper.Implementation;
using Microsoft.Extensions.Logging;
using NSubstitute;
using Xunit;
using Xunit.Abstractions;

namespace LittleBlocks.Excel.ClosedXml.IntegrationTests
{
    public class ExcelMapperPerformanceDataTests : IClassFixture<DatafileFixture>
    {
        public ExcelMapperPerformanceDataTests(DatafileFixture fixture, ITestOutputHelper output)
        {
            var fakeLoggerFactory = Substitute.For<ILoggerFactory>();
            fakeLoggerFactory.CreateLogger(Arg.Any<string>()).Returns(Substitute.For<ILogger>());
            
            var fakeWorkbookLoader = Substitute.For<IWorkbookLoader>();

            var factory = new ExcelMapperBuilder(fakeWorkbookLoader, fakeLoggerFactory);
            _fixture = fixture;
            _output = output;
            _sut = factory.Build(_fixture.GetWorkbook(Workbook));
        }

        private readonly DatafileFixture _fixture;
        private readonly ITestOutputHelper _output;
        private IExcelMapper _sut;

        private const string Workbook = "DataSheet_Person_1000_rows.xlsx";
        private const string SheetName = "Sheet1";

        [Fact]
        public void ShouldCorrectlyLoadFirstAndLastRows()
        {
            // Arrange
            // Act
            var result = _sut.Map<PersonModel>(SheetName, opt =>
                opt
                    .ForMember(m => m.CustomRowNumber, Resolve.ByValue("Custom Identification No"))).ToList();

            // Assert
            Assert.Equal(202, result.First().CustomRowNumber);
            Assert.Equal(1201, result.Last().CustomRowNumber);
        }

        [Fact]
        public void ShouldLoad1000RowsInUnder2Seconds()
        {
            // Arrange
            var sw = new Stopwatch();
            sw.Start();

            // Act
            var result = _sut.Map<PersonModel>(SheetName, opt =>
                opt
                    .ForMember(m => m.CustomRowNumber, Resolve.ByValue("Custom Identification No"))).ToList();

            sw.Stop();

            // Assert
            Assert.Equal(1000, result.Count);
            Assert.True(5000 > sw.ElapsedMilliseconds);

            _output.WriteLine("{0} records loaded in {1} milliseconds.", result.Count, sw.ElapsedMilliseconds);
        }

        [Fact]
        public void ShouldLoad1000RowsWithMetadataInUnder3Seconds()
        {
            // Arrange
            var sw = new Stopwatch();
            sw.Start();

            // Act
            var result = _sut.Map<PersonWithMetadataModel>(SheetName, opt =>
                opt
                    .ForMember(m => m.CustomRowNumber, Resolve.ByValue("Custom Identification No"))).ToList();

            sw.Stop();

            // Assert
            Assert.Equal(1000, result.Count);
            Assert.True(4000 > sw.ElapsedMilliseconds);

            _output.WriteLine("{0} records with metadata loaded in {1} milliseconds.", result.Count,
                sw.ElapsedMilliseconds);
        }
    }
}