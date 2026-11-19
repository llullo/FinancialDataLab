using System;
using FinancialDataLab.Application.Services;
using Xunit;

namespace FinancialDataLab.UnitTests.Services
{
    public class FinancialCalculationServiceTests
    {
        private readonly FinancialCalculationService _service =
            new FinancialCalculationService();

        [Theory]
        [InlineData(100, 125, 25)]
        [InlineData(200, 150, -25)]
        [InlineData(100, 100, 0)]
        [InlineData(80, 100, 25)]
        public void CalculatePercentageChange_ReturnsExpectedResult(
            int previousValue,
            int currentValue,
            int expected)
        {
            decimal result = _service.CalculatePercentageChange(
                previousValue,
                currentValue);

            Assert.Equal((decimal)expected, result);
        }

        [Fact]
        public void CalculatePercentageChange_WhenPreviousValueIsZero_Throws()
        {
            Assert.Throws<ArgumentException>(() =>
                _service.CalculatePercentageChange(0m, 100m));
        }
    }
}