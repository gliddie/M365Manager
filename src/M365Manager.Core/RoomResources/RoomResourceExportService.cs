using ClosedXML.Excel;

namespace M365Manager.Core.RoomResources;

/// <summary>Exports the (filtered) Rooms &amp; Resources overview grid to Excel.</summary>
public static class RoomResourceExportService
{
    public static void ExportToExcel(string filePath, IReadOnlyList<RoomResourceOverviewRow> rows)
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("RoomsAndResources");

        var headers = new[]
        {
            "Display name", "Type", "Alias", "Address", "Capacity", "Site", "Location", "Time zone",
            "Room list", "Delegate group", "Users group", "Access model", "Booking policy",
            "Booking window (days)", "Max duration (min)", "Recurring allowed", "Created", "Last imported",
        };
        for (var c = 0; c < headers.Length; c++)
            sheet.Cell(1, c + 1).Value = headers[c];
        sheet.Row(1).Style.Font.Bold = true;

        for (var r = 0; r < rows.Count; r++)
        {
            var row = rows[r];
            sheet.Cell(r + 2, 1).Value = row.DisplayName;
            sheet.Cell(r + 2, 2).Value = row.ResourceKind;
            sheet.Cell(r + 2, 3).Value = row.Alias;
            sheet.Cell(r + 2, 4).Value = row.PrimarySmtpAddress;
            sheet.Cell(r + 2, 5).Value = row.Capacity;
            sheet.Cell(r + 2, 6).Value = row.SiteCode;
            sheet.Cell(r + 2, 7).Value = row.Office;
            sheet.Cell(r + 2, 8).Value = row.TimeZone;
            sheet.Cell(r + 2, 9).Value = row.RoomListName;
            sheet.Cell(r + 2, 10).Value = row.DelegateGroup;
            sheet.Cell(r + 2, 11).Value = row.UsersGroup;
            sheet.Cell(r + 2, 12).Value = row.AccessModel;
            sheet.Cell(r + 2, 13).Value = row.BookingPolicy;
            sheet.Cell(r + 2, 14).Value = row.BookingWindowInDays;
            sheet.Cell(r + 2, 15).Value = row.MaximumDurationInMinutes;
            sheet.Cell(r + 2, 16).Value = BoolLabel(row.AllowRecurringMeetings);
            sheet.Cell(r + 2, 17).Value = row.CreatedDateTime;
            sheet.Cell(r + 2, 18).Value = row.LastImportedAtUtc;
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
