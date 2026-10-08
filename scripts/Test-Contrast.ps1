<#
.SYNOPSIS
    Measures WCAG contrast for the colour-token pairs the UI actually renders, in both palettes.
.DESCRIPTION
    The design brief asked for WCAG AA. Eyeballing a palette does not get you there - four pairs
    failed the first time this ran, including one where a code comment asserted a ratio that had
    been estimated rather than measured.

    Run it after changing anything in Themes/Colors.*.xaml. It reads the token values straight out
    of the dictionaries, so it cannot drift from what ships.

    Thresholds: 4.5 for normal text (WCAG 1.4.3 AA), 3.0 for UI component boundaries and large
    text (1.4.11 / 1.4.3). Exits 1 if any pair is below target.
.EXAMPLE
    pwsh scripts/Test-Contrast.ps1
#>
param(
    [string]$ThemesDir = (Join-Path (Split-Path -Parent $PSScriptRoot) "src/M365Manager/Themes")
)

function Get-Colors($path) {
    $map = @{}
    foreach ($line in Get-Content $path) {
        if ($line -match '<Color x:Key="(\w+)">#([0-9A-Fa-f]{6,8})</Color>') {
            $hex = $Matches[2]
            if ($hex.Length -eq 8) { $hex = $hex.Substring(2) }   # strip alpha
            $map[$Matches[1]] = $hex
        }
    }
    $map
}

function Get-Luminance($hex) {
    $ch = @(0,2,4) | ForEach-Object {
        $v = [Convert]::ToInt32($hex.Substring($_,2),16) / 255.0
        if ($v -le 0.03928) { $v / 12.92 } else { [Math]::Pow(($v + 0.055)/1.055, 2.4) }
    }
    0.2126*$ch[0] + 0.7152*$ch[1] + 0.0722*$ch[2]
}

function Get-Ratio($a, $b) {
    $la = Get-Luminance $a; $lb = Get-Luminance $b
    $hi = [Math]::Max($la,$lb); $lo = [Math]::Min($la,$lb)
    [Math]::Round(($hi + 0.05) / ($lo + 0.05), 2)
}

# fg, bg, label, minimum. 4.5 = AA normal text, 3.0 = AA large text / UI component boundary.
$pairs = @(
    @('TextColor','PageColor','body text on page',4.5),
    @('TextColor','SurfaceColor','body text on card',4.5),
    @('TextColor','SurfaceAltColor','body text on inset',4.5),
    @('TextMutedColor','PageColor','muted text on page',4.5),
    @('TextMutedColor','SurfaceColor','muted text on card',4.5),
    @('TextMutedColor','SurfaceAltColor','muted text on inset',4.5),
    @('SuccessColor','SuccessSoftColor','success pill',4.5),
    @('WarningColor','WarningSoftColor','warning pill',4.5),
    @('DangerColor','DangerSoftColor','danger pill',4.5),
    @('InfoColor','InfoSoftColor','info pill',4.5),
    @('DangerColor','SurfaceColor','danger button label',4.5),
    @('BorderStrongColor','SurfaceColor','input border',3.0),
    @('FocusColor','SurfaceColor','focus ring',3.0),
    @('ConsoleTextColor','ConsoleBackColor','console text',4.5),
    @('ConsoleMutedColor','ConsoleBackColor','console muted',4.5),
    @('BrandTextColor','SurfaceColor','brand TEXT on card',4.5),
    @('BrandTextColor','BrandSoftColor','brand TEXT on brand-soft',4.5),
    @('TextColor','BrandSoftColor','text on highlighted list entry / grid hover',4.5),
    @('TextColor','WarningSoftColor','session strip text (roles expiring)',4.5),
    @('TextColor','DangerSoftColor','session strip text (roles expired)',4.5),
    @('WarningColor','WarningSoftColor','session strip icon / border (warning)',3.0),
    @('DangerColor','DangerSoftColor','session strip icon / border (error)',3.0),
    @('BrandTextColor','SurfaceColor','selected-entry accent bar in drop-down',3.0),
    @('BrandTextColor','SidebarColor','nav accent on sidebar',3.0),
    @('OnBrandColor','BrandColor','white label on primary button',4.5),
    @('OnBrandColor','BrandHoverColor','white label on primary hover',4.5),
    @('OnBrandColor','DangerFillColor','white label on confirm button',4.5),
    @('DangerColor','SurfaceColor','confirm button outline',3.0)
)

$fail = 0
foreach ($theme in 'Light','Dark') {
    $c = Get-Colors (Join-Path $ThemesDir "Colors.$theme.xaml")
    Write-Output ""
    Write-Output "=== $theme ==="
    foreach ($p in $pairs) {
        if (-not $c.ContainsKey($p[0]) -or -not $c.ContainsKey($p[1])) {
            Write-Output ("  ?     {0} (key missing)" -f $p[2]); continue
        }
        $r = Get-Ratio $c[$p[0]] $c[$p[1]]
        $min = [double]$p[3]
        $ok = $r -ge $min
        if (-not $ok) { $fail++ }
        Write-Output ("  {0} {1,5:N2}  (min {2})  {3}" -f $(if($ok){'PASS'}else{'FAIL'}), $r, $min, $p[2])
    }
}
Write-Output ""
Write-Output "$fail pair(s) below target"
if ($fail -gt 0) { exit 1 }

