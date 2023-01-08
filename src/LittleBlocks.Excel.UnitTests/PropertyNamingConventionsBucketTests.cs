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
using LittleBlocks.Excel.Mapper;
using LittleBlocks.Excel.Mapper.PropertyMap.Conventions;
using Microsoft.Extensions.Logging;
using NSubstitute;
using Xunit;

namespace LittleBlocks.Excel.UnitTests
{
    public class PropertyNamingConventionsBucketTests
    {
        [Theory]
        [InlineData("Security Type", "SecurityType")]
        [InlineData("Issuer Industry Classification - Moody's", "IssuerIndustryClassificationMoodys")]
        [InlineData("Deal Issue (Derived) Rating - S&P", "DealIssueDerivedRatingSnP")]
        [InlineData("Issuer Rating Watchlist Status - Senior Unsecured - Moody's",
            "IssuerRatingWatchlistStatusSeniorUnsecuredMoodys")]
        [InlineData("Deal Issue (Derived) Rating - Moody's 2", "DealIssueDerivedRatingMoodys2")]
        public void ShouldCorrectlyRenameProperty(string sourceValue, string expectedValue)
        {
            // Arrange
            var sut = new PropertyNamingConventionsBucket(new List<IPropertyNameConvention>
            {
                new NameWithNoNumberAtStartConvention(),
                new ReplaceAmperstandWithNConvention(),
                new RemoveNonLetterAndDigitCharactersConvention()
            }, Substitute.For<ILogger<PropertyNamingConventionsBucket>>());

            // Act
            var result = sut.ApplyConvention(sourceValue);

            // Assert
            Assert.Equal(expectedValue, result);
        }

        [Fact]
        public void ShouldApplyAllConvetionsInGivenOrder()
        {
            // Arrange
            var firstConvention = Substitute.For<IPropertyNameConvention>();
            var secondConvention = Substitute.For<IPropertyNameConvention>();
            var sut = new PropertyNamingConventionsBucket(new List<IPropertyNameConvention>
            {
                firstConvention,
                secondConvention
            }, Substitute.For<ILogger<PropertyNamingConventionsBucket>>());

            // Act
            sut.ApplyConvention("PropertyName");

            // Assert
            Received.InOrder(() =>
            {
                firstConvention.ApplyConvention(Arg.Any<string>());
                secondConvention.ApplyConvention(Arg.Any<string>());
            });
        }

        [Fact]
        public void ShouldApplyAllConvetionsOnce()
        {
            // Arrange
            var firstConvention = Substitute.For<IPropertyNameConvention>();
            var secondConvention = Substitute.For<IPropertyNameConvention>();
            var sut = new PropertyNamingConventionsBucket(new List<IPropertyNameConvention>
            {
                firstConvention,
                secondConvention
            }, Substitute.For<ILogger<PropertyNamingConventionsBucket>>());

            // Act
            sut.ApplyConvention("PropertyName");

            // Assert
            firstConvention.Received(1).ApplyConvention(Arg.Any<string>());
            secondConvention.Received(1).ApplyConvention(Arg.Any<string>());
        }

        [Fact]
        public void ShouldNotThrowIfConvetionsCollectionIsEmpty()
        {
            // Arrange
            // Act
            var sut = new PropertyNamingConventionsBucket(new List<IPropertyNameConvention>(),
                Substitute.For<ILogger<PropertyNamingConventionsBucket>>());

            // Assert
            Assert.IsType<PropertyNamingConventionsBucket>(sut);
        }

        [Fact]
        public void ShouldReturnResultFromLastConvention()
        {
            // Arrange
            var firstConvention = Substitute.For<IPropertyNameConvention>();
            var secondConvention = Substitute.For<IPropertyNameConvention>();
            var sut = new PropertyNamingConventionsBucket(new List<IPropertyNameConvention>
            {
                firstConvention,
                secondConvention
            }, Substitute.For<ILogger<PropertyNamingConventionsBucket>>());
            secondConvention.ApplyConvention(Arg.Any<string>()).Returns("ModifiedPropertyName");

            // Act
            var result = sut.ApplyConvention("PropertyName");

            // Assert
            Assert.Equal("ModifiedPropertyName", result);
        }

        [Fact]
        public void ShouldReturnUnchangedValueIfNoPropertyNamingConvetionsFound()
        {
            // Arrange
            var sut = new PropertyNamingConventionsBucket(new List<IPropertyNameConvention>(),
                Substitute.For<ILogger<PropertyNamingConventionsBucket>>());

            // Act
            var result = sut.ApplyConvention("OriginalPropertyName");

            // Assert
            Assert.Equal("OriginalPropertyName", result);
        }

        [Fact]
        public void ShouldThrowIfNullPassInsteadIfConvetions()
        {
            // Arrange
            // Act
            Action action = () =>
                new PropertyNamingConventionsBucket(null, Substitute.For<ILogger<PropertyNamingConventionsBucket>>());

            // Assert
            Assert.Throws<ArgumentNullException>(action);
        }
    }
}