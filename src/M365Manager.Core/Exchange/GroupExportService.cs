using System.Text;
using ClosedXML.Excel;

namespace M365Manager.Core.Exchange;

/// <summary>Exports a group's details and member list to Excel or CSV.</summary>
public static class GroupExportService
{
    public static void ExportToExcel(string filePath, DistributionGroupInfo group, IReadOnlyList<GroupMemberInfo> members)
    {
        using var workbook = new XLWorkbook();

        var groupSheet = workbook.Worksheets.Add("Group");
        var groupRows = new (string Label, string Value)[]
        {
            ("Display name", group.DisplayName),
            ("Primary SMTP address", group.PrimarySmtpAddress),
            ("Alias", group.Alias),
            ("Type", group.GroupTypeLabel),
            ("Owner(s)", string.Join(", ", group.ManagedBy)),
            ("Notes", group.Notes),
            ("Status", group.IsDeletedInM365 ? "Deleted in M365" : "Active"),
        };
        for (var i = 0; i < groupRows.Length; i++)
        {
            groupSheet.Cell(i + 1, 1).Value = groupRows[i].Label;
            groupSheet.Cell(i + 1, 1).Style.Font.Bold = true;
            groupSheet.Cell(i + 1, 2).Value = groupRows[i].Value;
        }
        groupSheet.Columns().AdjustToContents();

        var memberSheet = workbook.Worksheets.Add("Members");
        var headers = new[] { "Display name", "Primary SMTP address", "User principal name", "Type", "Nested", "Path" };
        for (var c = 0; c < headers.Length; c++)
            memberSheet.Cell(1, c + 1).Value = headers[c];
        memberSheet.Row(1).Style.Font.Bold = true;

        for (var r = 0; r < members.Count; r++)
        {
            var m = members[r];
            memberSheet.Cell(r + 2, 1).Value = m.DisplayName;
            memberSheet.Cell(r + 2, 2).Value = m.PrimarySmtpAddress;
            memberSheet.Cell(r + 2, 3).Value = m.UserPrincipalName;
            memberSheet.Cell(r + 2, 4).Value = m.RecipientType;
            memberSheet.Cell(r + 2, 5).Value = m.IsNested ? "Yes" : "No";
            memberSheet.Cell(r + 2, 6).Value = m.Path;
        }
        if (members.Count > 0)
            memberSheet.RangeUsed()?.SetAutoFilter();
        memberSheet.Columns().AdjustToContents();

        workbook.SaveAs(filePath);
    }

    public static void ExportToCsv(string filePath, DistributionGroupInfo group, IReadOnlyList<GroupMemberInfo> members)
    {
        using var writer = new StreamWriter(filePath, false, new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));

        writer.WriteLine("Group details");
        WriteCsvRow(writer, "Display name", group.DisplayName);
        WriteCsvRow(writer, "Primary SMTP address", group.PrimarySmtpAddress);
        WriteCsvRow(writer, "Alias", group.Alias);
        WriteCsvRow(writer, "Type", group.GroupTypeLabel);
        WriteCsvRow(writer, "Owner(s)", string.Join("; ", group.ManagedBy));
        WriteCsvRow(writer, "Notes", group.Notes);
        WriteCsvRow(writer, "Status", group.IsDeletedInM365 ? "Deleted in M365" : "Active");
        writer.WriteLine();

        writer.WriteLine("Members");
        WriteCsvRow(writer, "Display name", "Primary SMTP address", "User principal name", "Type", "Nested", "Path");
        foreach (var m in members)
            WriteCsvRow(writer, m.DisplayName, m.PrimarySmtpAddress, m.UserPrincipalName, m.RecipientType, m.IsNested ? "Yes" : "No", m.Path);
    }

    private static void WriteCsvRow(TextWriter writer, params string?[] values)
        => writer.WriteLine(string.Join(",", values.Select(CsvEscape)));

    private static string CsvEscape(string? value)
    {
        value ??= "";
        return value.IndexOfAny([',', '"', '\n', '\r']) >= 0
            ? "\"" + value.Replace("\"", "\"\"") + "\""
            : value;
    }
}
