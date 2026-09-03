namespace M365Manager.Core.PowerShell;

/// <summary>How a transcript line is rendered in the console view.</summary>
public enum TranscriptLineKind
{
    /// <summary>The command about to run, shown after a prompt.</summary>
    Command,

    /// <summary>
    /// A Microsoft Graph request, shown after its own prompt. Graph is the other half of what the
    /// app does - a console showing only the PowerShell side made whole features (Trace, Teams
    /// policies, notification mail) look like they did nothing at all.
    /// </summary>
    GraphRequest,

    /// <summary>Whatever the command wrote to the pipeline, formatted the way PowerShell would print it.</summary>
    Output,

    /// <summary>The error stream.</summary>
    Error,

    /// <summary>The warning stream.</summary>
    Warning,

    /// <summary>Host/information stream - Write-Host and the like.</summary>
    Information,

    /// <summary>A separator naming the operation that follows.</summary>
    Header,
}

public sealed record TranscriptLine(TranscriptLineKind Kind, string Text, DateTime TimestampUtc);

/// <summary>
/// Collects the PowerShell commands and Graph requests the app makes, plus their output, so the UI
/// can show a console view of what is happening. Populated centrally - by
/// <see cref="PowerShellHost"/> for PowerShell and by
/// <see cref="M365Manager.Core.M365.GraphRestClient"/> for Graph - so no individual service has to
/// opt in.
///
/// Sign-in is deliberately excluded on both sides (see the suppressTranscript flag on
/// <see cref="PowerShellHost.InvokeAsync"/>, and M365AuthService's own SDK client) - those calls
/// carry device codes and tokens.
/// </summary>
public interface IPowerShellTranscript
{
    IReadOnlyList<TranscriptLine> Lines { get; }

    /// <summary>Raised on whichever thread produced the line - subscribers must marshal to the UI themselves.</summary>
    event EventHandler<TranscriptLine>? LineAdded;

    event EventHandler? Cleared;

    void Write(TranscriptLineKind kind, string text);

    /// <summary>Starts a visually separated block for one user-facing operation.</summary>
    void BeginOperation(string title);

    void Clear();

    /// <summary>The whole transcript as plain text, for the copy button.</summary>
    string GetText();
}
