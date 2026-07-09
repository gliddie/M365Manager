namespace M365Manager.Core.M365;

/// <summary>The M365 admin who signed in interactively (used for auditing/logging).</summary>
public sealed record SignedInUser(string Upn, string DisplayName, string Id);
