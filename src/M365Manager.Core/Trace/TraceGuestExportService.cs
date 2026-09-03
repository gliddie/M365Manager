using ClosedXML.Excel;

namespace M365Manager.Core.Trace;

/// <summary>Exports the (filtered) Trace guest list to Excel.</summary>
public static class TraceGuestExportService
{
    public static void ExportToExcel(string filePath, IReadOnlyList<TraceGuestRow> rows)
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("TraceGuests");

        var headers = new[]
        {
            "Display name", "E-mail", "User principal name", "Company", "Status",
            "Status changed (UTC)", "Created (UTC)", "Enabled",
        };
        for (var c = 0; c < headers.Length; c++)
            sheet.Cell(1, c + 1).Value = headers[c];
        sheet.Row(1).Style.Font.Bold = true;

        for (var r = 0; r < rows.Count; r++)
        {
            var row = rows[r];
            sheet.Cell(r + 2, 1).Value = row.DisplayName;
            sheet.Cell(r + 2, 2).Value = row.Mail;
            sheet.Cell(r + 2, 3).Value = row.UserPrincipalName;
            sheet.Cell(r + 2, 4).Value = row.CompanyName;
            sheet.Cell(r + 2, 5).Value = row.StatusLabel;
            sheet.Cell(r + 2, 6).Value = row.StateChangedUtc;
            sheet.Cell(r + 2, 7).Value = row.CreatedUtc;
            sheet.Cell(r + 2, 8).Value = row.AccountEnabled ? "Yes" : "No";
        }

        if (rows.Count > 0)
            sheet.RangeUsed()?.SetAutoFilter();
        sheet.Columns().AdjustToContents();

        workbook.SaveAs(filePath);
    }
}
