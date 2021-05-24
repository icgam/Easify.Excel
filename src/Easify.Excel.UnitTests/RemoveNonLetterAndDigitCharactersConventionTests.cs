// This software is part of the Easify.Excel Library
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
using Easify.Excel.Mapper.PropertyMap.Conventions;
using Xunit;

namespace Easify.Excel.UnitTests
{
    public class RemoveNonLetterAndDigitCharactersConventionTests : IDisposable
    {
        public RemoveNonLetterAndDigitCharactersConventionTests()
        {
            _sut = new RemoveNonLetterAndDigitCharactersConvention();
        }

        public void Dispose()
        {
            _sut = null;
        }

        private IPropertyNameConvention _sut;

        [Fact]
        public void ShouldRemoveAllSpecialSymbols()
        {
            // Arrange
            var sourceValue = "1report!\"£$%^ &*()    _ += -`¬;:'@#~/?.>   ,<\\|admin1";

            // Act
            var result = _sut.ApplyConvention(sourceValue);

            // Assert
            Assert.Equal("1reportadmin1", result);
        }
    }
}