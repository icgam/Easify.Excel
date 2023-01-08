// This software is part of the LittleBlocks.Ef Library
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
using System.IO;
using LittleBlocks.Excel.Exceptions;
using LittleBlocks.Testing;
using Xunit;

namespace LittleBlocks.Excel.ClosedXml.IntegrationTests
{
    public class ExcelMapperForCorruptFileTests
    {
        private const string Workbook = "SampleDataCorruptFile.xlsx";

        [Theory]
        [AutoSubstituteAndData]
        public void ShouldThrowIfWorkbookIsCorrupted(WorkbookLoader sut)
        {
            // Arrange
            var datafilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data", Workbook);

            // Act
            // Assert
            Assert.Throws<WorkbookIsCorruptedException>(() => sut.Load(datafilePath));
        }
    }
}