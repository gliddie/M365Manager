using ClosedXML.Excel;
using M365Manager.Data.Logging;

namespace M365Manager.Core.Logging;

/// <summary>Exports the (filtered) log view to Excel.</summary>
public static class LogExportService
{
    public static void ExportToExcel(string filePath, IReadOnlyList<LogEntry> rows)
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Logs");

        var headers = new[]
        {
            "Time (UTC)", "Severity", "Code", "Area", "Action", "Target", "Task#",
            "User", "Windows user", "Machine", "Message", "Correlation id",
        };
        for (var c = 0; c < headers.Length; c++)
            sheet.Cell(1, c + 1).Value = headers[c];
        sheet.Row(1).Style.Font.Bold = true;

        for (var r = 0; r < rows.Count; r++)
        {
            var row = rows[r];
            sheet.Cell(r + 2, 1).Value = row.TimestampUtc;
            sheet.Cell(r + 2, 1).Style.DateFormat.Format = "yyyy-mm-dd hh:mm:ss";
            sheet.Cell(r + 2, 2).Value = row.Severity.ToString();
            sheet.Cell(r + 2, 3).Value = row.EventCode;
            sheet.Cell(r + 2, 4).Value = row.Area;
            sheet.Cell(r + 2, 5).Value = row.Action;
            sheet.Cell(r + 2, 6).Value = row.TargetObject;
            sheet.Cell(r + 2, 7).Value = row.TaskNumber;
            sheet.Cell(r + 2, 8).Value = row.UserUpn;
            sheet.Cell(r + 2, 9).Value = row.WindowsUser;
            sheet.Cell(r + 2, 10).Value = row.MachineName;
            sheet.Cell(r + 2, 11).Value = row.Message;
            sheet.Cell(r + 2, 12).Value = row.CorrelationId?.ToString();
        }

        if (rows.Count > 0)
            sheet.RangeUsed()?.SetAutoFilter();

        sheet.Columns().AdjustToContents();
        // The message column carries whole snapshots (room removal), so cap it and wrap instead of
        // letting AdjustToContents produce a column thousands of characters wide.
        sheet.Column(11).Width = 120;
        sheet.Column(11).Style.Alignment.WrapText = true;

        workbook.SaveAs(filePath);
    }
}
