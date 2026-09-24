using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using M365Manager.Core.PowerShell;

namespace M365Manager.Controls;

/// <summary>
/// Console view over an <see cref="IPowerShellTranscript"/>, styled after the classic PowerShell
/// host. A RichTextBox rather than a list, so the text stays selectable and Ctrl+C works while
/// each stream still gets its own colour.
/// </summary>
public partial class PowerShellConsole : UserControl
{
    /// <summary>The transcript to display. Set from XAML, usually bound to the page's ViewModel.</summary>
    public static readonly DependencyProperty TranscriptProperty = DependencyProperty.Register(
        nameof(Transcript), typeof(IPowerShellTranscript), typeof(PowerShellConsole),
        new PropertyMetadata(null, OnTranscriptChanged));

    public IPowerShellTranscript? Transcript
    {
        get => (IPowerShellTranscript?)GetValue(TranscriptProperty);
        set => SetValue(TranscriptProperty, value);
    }

    private readonly Paragraph _paragraph = new();

    public PowerShellConsole()
    {
        InitializeComponent();

        // One paragraph holding Runs, so line spacing matches a real console instead of
        // paragraph spacing.
        Console.Document.Blocks.Clear();
        Console.Document.Blocks.Add(_paragraph);
        Console.Document.PageWidth = 2000; // suppress word wrap; the horizontal scrollbar handles it

        Unloaded += (_, _) => Detach(Transcript);
    }

    private static void OnTranscriptChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        var console = (PowerShellConsole)d;
        console.Detach(e.OldValue as IPowerShellTranscript);
        console.Attach(e.NewValue as IPowerShellTranscript);
    }

    private void Attach(IPowerShellTranscript? transcript)
    {
        if (transcript is null)
            return;

        transcript.LineAdded += OnLineAdded;
        transcript.Cleared += OnCleared;
        Rebuild(transcript);
    }

    private void Detach(IPowerShellTranscript? transcript)
    {
        if (transcript is null)
            return;

        transcript.LineAdded -= OnLineAdded;
        transcript.Cleared -= OnCleared;
    }

    // Both handlers arrive on the PowerShell worker thread.
    private void OnLineAdded(object? sender, TranscriptLine line)
        => Dispatcher.BeginInvoke(() => Append(line));

    private void OnCleared(object? sender, EventArgs e)
        => Dispatcher.BeginInvoke(() => Rebuild(Transcript));

    private void Rebuild(IPowerShellTranscript? transcript)
    {
        _paragraph.Inlines.Clear();
        if (transcript is not null)
        {
            foreach (var line in transcript.Lines)
                Append(line, scroll: false);
        }
        ScrollIfWanted();
        UpdateStatus();
    }

    private void Append(TranscriptLine line, bool scroll = true)
    {
        // A command gets the familiar prompt in front of it; a Graph request gets its own, so the
        // two halves of an operation stay distinguishable in one console.
        if (line.Kind == TranscriptLineKind.Command)
            _paragraph.Inlines.Add(new Run("PS> ") { Foreground = Brushes.White, FontWeight = FontWeights.Bold });
        else if (line.Kind == TranscriptLineKind.GraphRequest)
            _paragraph.Inlines.Add(new Run("GRAPH> ") { Foreground = BrushFor(line.Kind), FontWeight = FontWeights.Bold });

        _paragraph.Inlines.Add(new Run(line.Text + Environment.NewLine)
        {
            Foreground = BrushFor(line.Kind),
            FontWeight = line.Kind is TranscriptLineKind.Command or TranscriptLineKind.GraphRequest or TranscriptLineKind.Header
                ? FontWeights.Bold
                : FontWeights.Normal,
        });

        if (scroll)
        {
            ScrollIfWanted();
            UpdateStatus();
        }
    }

    /// <summary>Stream colours as the PowerShell host uses them.</summary>
    private static Brush BrushFor(TranscriptLineKind kind) => kind switch
    {
        TranscriptLineKind.Command => Brushes.White,
        // Graph's own blue-violet, so a REST call is not mistaken for a cmdlet at a glance.
        TranscriptLineKind.GraphRequest => new SolidColorBrush(Color.FromRgb(0xC5, 0x86, 0xC0)),
        TranscriptLineKind.Error => new SolidColorBrush(Color.FromRgb(0xFF, 0x60, 0x60)),
        TranscriptLineKind.Warning => new SolidColorBrush(Color.FromRgb(0xF5, 0xD7, 0x6E)),
        TranscriptLineKind.Information => new SolidColorBrush(Color.FromRgb(0x9C, 0xDC, 0xFE)),
        TranscriptLineKind.Header => new SolidColorBrush(Color.FromRgb(0x7E, 0xD3, 0x21)),
        _ => new SolidColorBrush(Color.FromRgb(0xDC, 0xDC, 0xDC)),
    };

    private void ScrollIfWanted()
    {
        if (AutoScrollBox.IsChecked == true)
            Console.ScrollToEnd();
    }

    private void UpdateStatus()
    {
        var count = Transcript?.Lines.Count ?? 0;
        StatusText.Text = count == 0 ? "" : $"{count} line(s)";
    }

    private void CopyButton_Click(object sender, RoutedEventArgs e)
    {
        var text = Transcript?.GetText() ?? "";
        if (text.Length == 0)
        {
            StatusText.Text = "Nothing to copy yet";
            return;
        }

        try
        {
            Clipboard.SetText(text);
            StatusText.Text = "Copied to clipboard";
        }
        catch
        {
            // The clipboard can be locked by another app.
            StatusText.Text = "Could not access the clipboard";
        }
    }

    private void ClearButton_Click(object sender, RoutedEventArgs e) => Transcript?.Clear();
}
