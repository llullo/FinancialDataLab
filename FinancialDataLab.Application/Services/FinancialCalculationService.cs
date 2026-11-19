using System;

namespace FinancialDataLab.Application.Services
{
    public class FinancialCalculationService
    {
        public decimal CalculatePercentageChange(
            decimal previousValue,
            decimal currentValue)
        {
            if (previousValue == 0m)
            {
                throw new ArgumentException(
                    "Previous value cannot be zero.",
                    nameof(previousValue));
            }

            return ((currentValue - previousValue) / previousValue) * 100m;
        }
    }
}