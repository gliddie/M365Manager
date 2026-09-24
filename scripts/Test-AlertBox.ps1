<#
.SYNOPSIS
    Checks that AlertBox honours an externally bound Visibility.
.DESCRIPTION
    AlertBox collapses itself when it has no text. It used to do that by assigning Visibility in
    code, which writes a LOCAL value - and in WPF a local value beats a binding the consumer set
    in XAML. That silently killed a "show only on conflict" binding on the Teams Policies page and
    left a warning about removing groups permanently on screen, which is alarming and wrong.

    The ordering matters: XAML establishes the Visibility binding first and assigns Text after, so
    the Text change handler was the thing that clobbered it. A test that binds after setting Text
    passes either way and proves nothing - both orders are checked here.

    Run after touching AlertBox, or any view that binds its Visibility.
.EXAMPLE
    pwsh scripts/Test-AlertBox.ps1
#>
param(
    [string]$AppDir = (Join-Path (Split-Path -Parent $PSScriptRoot) "src/M365Manager/bin/Debug/net10.0-windows"),
    [string]$ThemesDir = (Join-Path (Split-Path -Parent $PSScriptRoot) "src/M365Manager/Themes")
)

$body = @'
$out = New-Object System.Collections.Generic.List[string]
try {
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
    [Reflection.Assembly]::LoadFrom((Join-Path $AppDir 'M365Manager.dll')) | Out-Null

    # Resources come off disk rather than through pack:// - LoadFrom does not register the
    # assembly for pack URI resolution, and the control resolves StaticResource at construction.
    $app = New-Object System.Windows.Application
    foreach ($n in 'Tokens.xaml','Colors.Light.xaml','Styles.xaml') {
        $fs = [System.IO.File]::OpenRead((Join-Path $ThemesDir $n))
        try { $app.Resources.MergedDictionaries.Add([System.Windows.Markup.XamlReader]::Load($fs)) }
        finally { $fs.Dispose() }
    }

    function Check($label, $box, $want) {
        $box.Measure([System.Windows.Size]::new(400,400))
        $got = $box.Visibility
        $mark = if ("$got" -eq $want) { 'PASS' } else { 'FAIL' }
        $out.Add(("  {0}  {1,-30} -> {2,-9} (want {3})" -f $mark, $label, $got, $want))
    }

    function BindVisibility($box, $flag) {
        $src = New-Object psobject -Property @{ Flag = $flag }
        $b = New-Object System.Windows.Data.Binding 'Flag'
        $b.Source = $src
        $b.Converter = New-Object M365Manager.Converters.BoolToVisibilityConverter
        [void][System.Windows.Data.BindingOperations]::SetBinding(
            $box, [System.Windows.UIElement]::VisibilityProperty, $b)
    }

    $a = New-Object M365Manager.Controls.AlertBox; $a.Text = ''
    Check 'empty text, no binding' $a 'Collapsed'

    $b = New-Object M365Manager.Controls.AlertBox; $b.Text = 'something happened'
    Check 'text set, no binding' $b 'Visible'

    # The regression: before the fix the control wrote Visibility itself, so this came out Visible
    # and a conditional warning showed unconditionally.
    # Order matters, and the real XAML order is what must be reproduced: the Visibility binding
    # is established first, then Text is assigned. If the control writes Visibility from inside
    # the Text change handler, that local value wipes the binding out - which is the bug.
    $c = New-Object M365Manager.Controls.AlertBox
    BindVisibility $c $false
    $c.Text = "in more than one group"
    Check "bind THEN set text, flag=false" $c "Collapsed"

    $d = New-Object M365Manager.Controls.AlertBox
    BindVisibility $d $true
    $d.Text = "in more than one group"
    Check "bind THEN set text, flag=true" $d "Visible"

    # The other order, for completeness - a binding applied after a local value replaces it, so
    # this case passed even with the broken control and proves nothing on its own.
    $e = New-Object M365Manager.Controls.AlertBox; $e.Text = "in more than one group"
    BindVisibility $e $false
    Check "set text THEN bind, flag=false" $e "Collapsed"
}
catch { $out.Add("  ERROR: $($_.Exception.Message)") }
$out
'@

$rs = [RunspaceFactory]::CreateRunspace()
$rs.ApartmentState = 'STA'
$rs.ThreadOptions = 'ReuseThread'
$rs.Open()
$rs.SessionStateProxy.SetVariable('AppDir', $AppDir)
$rs.SessionStateProxy.SetVariable('ThemesDir', $ThemesDir)

$ps = [PowerShell]::Create()
$ps.Runspace = $rs
[void]$ps.AddScript($body)
$out = $ps.Invoke()
$out | ForEach-Object { $_ }
$ps.Streams.Error | ForEach-Object { "  STREAM ERROR: $_" }
$ps.Dispose(); $rs.Close()

if (($out | Where-Object { $_ -match 'FAIL|ERROR' })) { exit 1 }
