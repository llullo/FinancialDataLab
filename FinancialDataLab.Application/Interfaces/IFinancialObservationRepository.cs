using System;
using System.Collections.Generic;
using FinancialDataLab.Application.Models;

namespace FinancialDataLab.Application.Interfaces
{
    public interface IFinancialObservationRepository
    {
        IReadOnlyList<FinancialObservation> GetByIndicator(
            string indicatorCode,
            DateTime fromDate,
            DateTime toDate);
        IReadOnlyList<FinancialObservation> GetNextPage(
            string indicatorCode,
            DateTime fromDate,
            DateTime toDate,
            int pageSize,
            DateTime? lastObservationDate,
            int? lastId);
        IReadOnlyList<FinancialObservation> GetPage(
            string indicatorCode,
            DateTime fromDate,
            DateTime toDate,
            int pageNumber,
            int pageSize);
    }
}