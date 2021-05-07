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


using System;
using Easify.Excel.Exceptions;
using Easify.Testing;
using Xunit;

namespace Easify.Excel.ClosedXml.IntegrationTests
{
    public class ExcelMapperForCorruptFileTests
    {
        private const string Workbook = "SampleDataCorruptFile.xlsx";

        [Theory]
        [AutoSubstituteAndData]
        public void ShouldThrowIfWorkbookIsCorrupted(WorkbookLoader sut)
        {
            // Arrange
            var datafilePath = $"{AppDomain.CurrentDomain.BaseDirectory}\\data\\{Workbook}";

            // Act
            // Assert
            Assert.Throws<WorkbookIsCorruptedException>(() => sut.Load(datafilePath));
        }
    }
}