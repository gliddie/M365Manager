using System.Text;
using System.IO;

namespace M365Manager.Core.PowerShell;

/// <summary>
/// A TextWriter that forwards each completed line to a callback. Used to capture
/// console output (e.g. the MSAL device-code sign-in prompt) that hosted PowerShell
/// / EXO writes via Console.WriteLine, which a GUI app would otherwise never see.
/// </summary>
internal sealed class ForwardingTextWriter : TextWriter
{
    private readonly Action<string> _onLine;
    private readonly StringBuilder _buffer = new();
    private readonly object _lock = new();

    public ForwardingTextWriter(Action<string> onLine) => _onLine = onLine;

    public override Encoding Encoding => Encoding.UTF8;

    public override void Write(char value)
    {
        lock (_lock)
        {
            if (value == '\n')
                FlushLine();
            else if (value != '\r')
                _buffer.Append(value);
        }
    }

    public override void Write(string? value)
    {
        if (string.IsNullOrEmpty(value))
            return;

        foreach (var c in value)
            Write(c);
    }

    private void FlushLine()
    {
        if (_buffer.Length == 0)
            return;

        var line = _buffer.ToString();
        _buffer.Clear();

        if (!string.IsNullOrWhiteSpace(line))
            _onLine(line);
    }

    protected override void Dispose(bool disposing)
    {
        lock (_lock)
            FlushLine();
        base.Dispose(disposing);
    }
}
