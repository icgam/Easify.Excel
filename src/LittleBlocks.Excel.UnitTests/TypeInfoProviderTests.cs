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

using System.Linq;
using LittleBlocks.Excel.Mapper.PropertyMap;
using LittleBlocks.Excel.UnitTests.Models;
using Xunit;

namespace LittleBlocks.Excel.UnitTests
{
    public class TypeInfoProviderTests
    {
        [Fact]
        public void ShouldReturnAllPublicPropertiesForGivenModel()
        {
            // Arrange
            var sut = new TypeInfoProvider();
            // Act
            var properties = sut.GetPublicProperties<Employee>().ToList();

            // Assert
            Assert.Equal(4, properties.Count);
            Assert.Contains(properties, p => p.Name == nameof(Employee.FirstName));
            Assert.Contains(properties, p => p.Name == nameof(Employee.LastName));
            Assert.Contains(properties, p => p.Name == nameof(Employee.Age));
            Assert.Contains(properties, p => p.Name == nameof(Employee.Salary));
        }
    }
}