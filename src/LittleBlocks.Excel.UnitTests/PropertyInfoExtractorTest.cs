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
using System.Linq.Expressions;
using LittleBlocks.Excel.Reflection;
using LittleBlocks.Excel.UnitTests.Helpers;
using LittleBlocks.Testing;
using Xunit;

namespace LittleBlocks.Excel.UnitTests
{
    public class PropertyInfoExtractorTest
    {
        [Theory]
        [AutoSubstituteAndData]
        public void ShouldReturnPropertyNameForFirstName(Person person)
        {
            // Arrange
            // Act
            var result = person.GetPropertyInfo(o => o.FirstName);

            // Assert
            Assert.Equal(nameof(Person.FirstName), result.Name);
        }

        [Theory]
        [AutoSubstituteAndData]
        public void ShouldReturnPropertyValueForFirstName(Person person)
        {
            // Arrange
            // Act
            var result = person.GetPropertyInfo(o => o.FirstName);

            // Assert
            Assert.Equal(person.FirstName, result.GetValue(person));
        }

        [Fact]
        public void ShouldReturnPropertyNameForFirstNameDirectlyOfExpression()
        {
            // Arrange
            Expression<Func<Person, object>> property = p => p.FirstName;
            // Act
            var result = property.GetName();

            // Assert
            Assert.Equal(nameof(Person.FirstName), result);
        }

        [Fact]
        public void ShouldReturnPropertyTypeNameForAgeIgnoringNullableWrapperDirectlyOfExpression()
        {
            // Arrange
            Expression<Func<Person, object>> property = p => p.Age;
            // Act
            var result = property.GetTypeName();

            // Assert
            Assert.Equal("Int32", result);
        }

        [Fact]
        public void ShouldReturnPropertyTypeNameForFirstNameDirectlyOfExpression()
        {
            // Arrange
            Expression<Func<Person, object>> property = p => p.FirstName;
            // Act
            var result = property.GetTypeName();

            // Assert
            Assert.Equal("String", result);
        }

        [Fact]
        public void ShouldReturnPropertyValueForFirstNameDirectlyOfExpression()
        {
            // Arrange
            var person = new Person
            {
                FirstName = "Jack"
            };
            Expression<Func<Person, object>> property = p => p.FirstName;
            // Act
            var result = property.GetValue(person);

            // Assert
            Assert.Equal("Jack", result);
        }
    }
}