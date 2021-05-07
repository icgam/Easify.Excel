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
using Easify.Excel.Mapper;
using Easify.Excel.Mapper.PropertyMap.Convetions;
using Xunit;

namespace Easify.Excel.UnitTests
{
    public class ReplaceAmperstandWithNConventionTests : IDisposable
    {
        public ReplaceAmperstandWithNConventionTests()
        {
            _sut = new ReplaceAmperstandWithNConvention();
        }

        public void Dispose()
        {
            _sut = null;
        }

        private IPropertyNameConvention _sut;

        [Fact]
        public void ShouldReturnAmendedValueWithNoAmperstand()
        {
            // Arrange
            // Act
            var result = _sut.ApplyConvention("Bower&Wilkins!");

            // Assert
            Assert.Equal("BowernWilkins!", result);
        }

        [Fact]
        public void ShouldReturnSameValueAsNoAmperstandsAreFound()
        {
            // Arrange
            // Act
            var result = _sut.ApplyConvention("BowerAndWilkins!");

            // Assert
            Assert.Equal("BowerAndWilkins!", result);
        }
    }
}