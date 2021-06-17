using ClosedXML.Excel.Drawings;
using Easify.Excel.ClosedXml;
using FluentAssertions;
using NSubstitute;
using Xunit;

namespace Easify.Excel.UnitTests.ClosedXml
{
    public class DataSheetPictureTests
    {
        [Theory]
        [InlineData(250, 50)]
        [InlineData(100, 20)]
        [InlineData(500, 100)]
        [InlineData(200, 40)]
        [InlineData(180, 36)]
        public void Should_WithWidth_CalculateHeightCorrectly(int width, int expectedHeight)
        {
            // Arrange
            var fakePicture = Substitute.For<IXLPicture>();
            fakePicture.Height.Returns(100);
            fakePicture.Width.Returns(500);

            var sut = new DataSheetPicture(fakePicture);
            
            // Act
            var actual = sut.WithWidth(width);

            // Assert
            actual.Should().NotBeNull();
            fakePicture.Received(1).WithSize(width, expectedHeight);
        }        
        
        [Theory]
        [InlineData(50 , 250)]
        [InlineData(20, 100)]
        [InlineData(100, 500)]
        [InlineData(40, 200)]
        [InlineData(36, 180)]
        public void Should_WithHeight_CalculateWidthCorrectly(int height, int expectedWidth)
        {
            // Arrange
            var fakePicture = Substitute.For<IXLPicture>();
            fakePicture.Height.Returns(100);
            fakePicture.Width.Returns(500);

            var sut = new DataSheetPicture(fakePicture);
            
            // Act
            var actual = sut.WithHeight(height);

            // Assert
            actual.Should().NotBeNull();
            fakePicture.Received(1).WithSize(expectedWidth, height);
        }
    }
}