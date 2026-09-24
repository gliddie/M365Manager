using System.Text;

namespace M365Manager.Core.PowerShell;

/// <summary>
/// In-memory <see cref="IPowerShellTranscript"/>. One instance for the whole app: the runspace is
/// shared and access to it is serialized, so a single console view genuinely reflects what is
/// running - there is never a second operation interleaving with it.
/// </summary>
public sealed class PowerShellTranscript : IPowerShellTranscript
{
    /// <summary>
    /// Guards against a runaway transcript. Only reached if someone works for hours without ever
    /// clearing; the oldest lines are dropped, which is why the view re-syncs from
    /// <see cref="Lines"/> rather than only appending.
    /// </summary>
    private const int MaxLines = 20_000;

    private readonly object _gate = new();
    private readonly List<TranscriptLine> _lines = new();

    public IReadOnlyList<TranscriptLine> Lines
    {
        get
        {
            lock (_gate)
                return _lines.ToArray();
        }
    }

    public event EventHandler<TranscriptLine>? LineAdded;
    public event EventHandler? Cleared;

    public void Write(TranscriptLineKind kind, string text)
    {
        if (text is null)
            return;

        // Multi-line output becomes one entry per line so the view can colour and virtualize them.
        foreach (var raw in text.Replace("\r\n", "\n").Split('\n'))
        {
            var line = new TranscriptLine(kind, raw.TrimEnd(), DateTime.UtcNow);

            bool overflowed;
            lock (_gate)
            {
                _lines.Add(line);
                overflowed = _lines.Count > MaxLines;
                if (overflowed)
                    _lines.RemoveRange(0, _lines.Count - MaxLines);
            }

            if (overflowed)
                Cleared?.Invoke(this, EventArgs.Empty); // makes the view rebuild from Lines
            else
                LineAdded?.Invoke(this, line);
        }
    }

    public void BeginOperation(string title)
    {
        Write(TranscriptLineKind.Header, "");
        Write(TranscriptLineKind.Header, $"=== {title} — {DateTime.Now:yyyy-MM-dd HH:mm:ss} ===");
    }

    public void Clear()
    {
        lock (_gate)
            _lines.Clear();

        Cleared?.Invoke(this, EventArgs.Empty);
    }

    public string GetText()
    {
        var sb = new StringBuilder();
        foreach (var line in Lines)
        {
            // Errors and warnings get their stream marked, so a pasted transcript still reads
            // correctly once the colours are gone.
            var prefix = line.Kind switch
            {
                TranscriptLineKind.Command => "PS> ",
                TranscriptLineKind.GraphRequest => "GRAPH> ",
                TranscriptLineKind.Error => "ERROR: ",
                TranscriptLineKind.Warning => "WARNING: ",
                _ => "",
            };
            sb.AppendLine(prefix + line.Text);
        }
        return sb.ToString();
    }
}
