
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

        // OFFSET/FETCH: retrieves a numbered page.
        IReadOnlyList<FinancialObservation> GetPage(
            string indicatorCode,
            DateTime fromDate,
            DateTime toDate,
            int pageNumber,
            int pageSize);

        // Keyset: continues from the last known observation.
        IReadOnlyList<FinancialObservation> GetNextPage(
            string indicatorCode,
            DateTime fromDate,
            DateTime toDate,
            int pageSize,
            DateTime? lastObservationDate,
            int? lastId);
    }
}
