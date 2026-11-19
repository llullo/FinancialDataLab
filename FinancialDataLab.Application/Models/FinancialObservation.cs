using System;

namespace FinancialDataLab.Application.Models
{
    public class FinancialObservation
    {
        public int Id { get; set; }

        public string IndicatorCode { get; set; }

        public DateTime ObservationDate { get; set; }

        public decimal Value { get; set; }
    }
}