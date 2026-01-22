using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;

namespace ProductCheckerV2.Common
{
    public class ExcelService
    {
        public List<ProductListingData> ReadExcelFile(string filePath)
        {
            var data = new List<ProductListingData>();

            try
            {
                using var workbook = new XLWorkbook(filePath);
                var worksheet = workbook.Worksheets.First();

                // Find the first row with data
                var firstRow = worksheet.FirstRowUsed();
                if (firstRow == null)
                    return data;

                // Try to detect header row
                int startRow = 1;
                for (int row = 1; row <= 5; row++)
                {
                    var cell1 = worksheet.Cell(row, 1).GetString().ToLower();
                    var cell2 = worksheet.Cell(row, 2).GetString().ToLower();
                    var cell3 = worksheet.Cell(row, 3).GetString().ToLower();

                    if ((cell1.Contains("listing") || cell1.Contains("id")) &&
                        (cell2.Contains("case") || cell2.Contains("number")) &&
                        (cell3.Contains("url") || cell3.Contains("link")))
                    {
                        startRow = row + 1;
                        break;
                    }
                }

                var lastRow = worksheet.LastRowUsed();
                for (int row = startRow; row <= lastRow.RowNumber(); row++)
                {
                    var listingId = worksheet.Cell(row, 1).GetString().Trim();
                    var caseNumber = worksheet.Cell(row, 2).GetString().Trim();
                    var url = worksheet.Cell(row, 3).GetString().Trim();

                    // Skip empty rows
                    if (string.IsNullOrWhiteSpace(listingId) &&
                        string.IsNullOrWhiteSpace(caseNumber) &&
                        string.IsNullOrWhiteSpace(url))
                        continue;

                    data.Add(new ProductListingData
                    {
                        ListingId = listingId ?? "",
                        CaseNumber = string.IsNullOrWhiteSpace(caseNumber) ? null : caseNumber,
                        ProductUrl = url ?? ""
                    });
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error reading Excel file: {ex.Message}", ex);
            }

            return data;
        }
    }

    public class ProductListingData
    {
        public string ListingId { get; set; }
        public string CaseNumber { get; set; } // Can be null
        public string ProductUrl { get; set; }
    }
}