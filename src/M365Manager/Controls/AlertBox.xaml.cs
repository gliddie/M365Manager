using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace M365Manager.Controls;

/// <summary>How an <see cref="AlertBox"/> reads. Drives both its colour and its glyph.</summary>
public enum AlertSeverity
{
    Info,
    Success,
    Warning,
    Error,
}

/// <summary>
/// One status message with a severity. Collapses itself when <see cref="Text"/> is empty, so a
/// view can bind it unconditionally instead of pairing it with a visibility converter.
/// </summary>
public partial class AlertBox : UserControl
{
    public AlertBox()
    {
        InitializeComponent();
        Apply();
    }

    public static readonly DependencyProperty TextProperty = DependencyProperty.Register(
        nameof(Text), typeof(string), typeof(AlertBox),
        new PropertyMetadata("", OnChanged));

    public static readonly DependencyProperty SeverityProperty = DependencyProperty.Register(
        nameof(Severity), typeof(AlertSeverity), typeof(AlertBox),
        new PropertyMetadata(AlertSeverity.Info, OnChanged));

    public string Text
    {
        get => (string)GetValue(TextProperty);
        set => SetValue(TextProperty, value);
    }

    public AlertSeverity Severity
    {
        get => (AlertSeverity)GetValue(SeverityProperty);
        set => SetValue(SeverityProperty, value);
    }

    private static void OnChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        => ((AlertBox)d).Apply();

    private void Apply()
    {
        // Visibility is deliberately NOT set here. Assigning it in code writes a local value, and
        // in WPF a local value beats a binding the consumer put in XAML - which silently killed a
        // "show this only on conflict" binding and left a scary warning permanently on screen.
        // The empty-text collapse now lives in a style trigger (AlertBox.xaml), which yields to an
        // external binding the way it should.

        // Glyphs are Segoe MDL2: Completed, Warning, ErrorBadge, Info.
        var (foreground, background, glyph) = Severity switch
        {
            AlertSeverity.Success => ("SuccessBrush", "SuccessSoftBrush", ""),
            AlertSeverity.Warning => ("WarningBrush", "WarningSoftBrush", ""),
            AlertSeverity.Error => ("DangerBrush", "DangerSoftBrush", ""),
            _ => ("InfoBrush", "InfoSoftBrush", ""),
        };

        Glyph.Text = glyph;
        Glyph.SetResourceReference(TextBlock.ForegroundProperty, foreground);
        Shell.SetResourceReference(BackgroundProperty, background);
        Shell.SetResourceReference(Border.BorderBrushProperty, foreground);
    }
}


