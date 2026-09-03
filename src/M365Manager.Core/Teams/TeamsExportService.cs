using ClosedXML.Excel;

namespace M365Manager.Core.Teams;

/// <summary>Exports the (filtered) Teams overview grid to Excel.</summary>
public static class TeamsExportService
{
    public static void ExportToExcel(string filePath, IReadOnlyList<TeamOverviewRow> rows)
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Teams");

        var headers = new[]
        {
            "Display name", "Alias", "Primary SMTP address", "Owner(s)", "Visibility", "Archived",
            "Guests allowed", "Hidden from address book", "Welcome message enabled",
            "SharePoint site URL", "Members", "Deleted in M365", "Created", "Last imported",
        };
        for (var c = 0; c < headers.Length; c++)
            sheet.Cell(1, c + 1).Value = headers[c];
        sheet.Row(1).Style.Font.Bold = true;

        for (var r = 0; r < rows.Count; r++)
        {
            var row = rows[r];
            sheet.Cell(r + 2, 1).Value = row.DisplayName;
            sheet.Cell(r + 2, 2).Value = row.Alias;
            sheet.Cell(r + 2, 3).Value = row.PrimarySmtpAddress;
            sheet.Cell(r + 2, 4).Value = row.Owners;
            sheet.Cell(r + 2, 5).Value = row.Visibility;
            sheet.Cell(r + 2, 6).Value = row.IsArchived ? "Yes" : "No";
            sheet.Cell(r + 2, 7).Value = BoolLabel(row.AllowToAddGuests);
            sheet.Cell(r + 2, 8).Value = BoolLabel(row.HiddenFromAddressListsEnabled);
            sheet.Cell(r + 2, 9).Value = BoolLabel(row.WelcomeMessageEnabled);
            sheet.Cell(r + 2, 10).Value = row.SharePointSiteUrl;
            sheet.Cell(r + 2, 11).Value = row.MemberCount;
            sheet.Cell(r + 2, 12).Value = row.IsDeletedInM365 ? "Yes" : "No";
            sheet.Cell(r + 2, 13).Value = row.CreatedDateTime;
            sheet.Cell(r + 2, 14).Value = row.LastImportedAtUtc;
        }

        if (rows.Count > 0)
            sheet.RangeUsed()?.SetAutoFilter();
        sheet.Columns().AdjustToContents();

        workbook.SaveAs(filePath);
    }

    private static string BoolLabel(bool? value) => value switch
    {
        true => "Yes",
        false => "No",
        null => "Unknown",
    };
}
