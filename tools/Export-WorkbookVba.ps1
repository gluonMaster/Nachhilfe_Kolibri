<#
.SYNOPSIS
Exports the VBA project of the Nachhilfe workbook into vba/.

.DESCRIPTION
Copies the workbook to a temporary folder, opens the copy read-only in a hidden
Excel instance with macros and events disabled, and writes every component into
the repository layout:

  standard modules   vba/modules/<name>.bas
  class modules      vba/classes/<name>.cls
  user forms         vba/forms/<name>.frm and .frx
  worksheet code     vba/worksheets/<sheet name>.vba   (export header removed)
  workbook code      vba/workbook/ThisWorkbook.vba     (export header removed)

Document modules without code are skipped. Existing files of these types are
replaced, so components removed from the workbook disappear from vba/ as well.
Code text is kept as exported; only line endings change to LF (CRLF for .frm).
Files are written as UTF-8: exported bytes that are already valid UTF-8 are
kept, anything else is decoded with the Windows ANSI code page.

If the Excel instance shows a dialog or does not finish within -TimeoutSec, the
script reports the dialog text, ends that Excel instance and fails without
touching vba/.

Requires desktop Excel with "Trust access to the VBA project object model".

.EXAMPLE
powershell -ExecutionPolicy Bypass -File tools\Export-WorkbookVba.ps1
powershell -ExecutionPolicy Bypass -File tools\Export-WorkbookVba.ps1 -WorkbookPath .\Nachhilfe_26-27.xlsm
#>
param(
    [string]$WorkbookPath,
    [int]$TimeoutSec = 120
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$vbaRoot = Join-Path $repoRoot 'vba'

if (-not $WorkbookPath) {
    $candidates = @(Get-ChildItem -LiteralPath $repoRoot -Filter 'Nachhilfe_*.xlsm' -File)
    if ($candidates.Count -ne 1) {
        throw "Pass -WorkbookPath: found $($candidates.Count) Nachhilfe_*.xlsm files in $repoRoot."
    }
    $WorkbookPath = $candidates[0].FullName
}
$WorkbookPath = (Resolve-Path -LiteralPath $WorkbookPath).Path

Add-Type @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

public static class ExportWatch {
    delegate bool EnumProc(IntPtr hwnd, IntPtr lParam);
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr lParam);
    [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr parent, EnumProc cb, IntPtr lParam);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr hwnd, StringBuilder s, int n);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(IntPtr hwnd, StringBuilder s, int n);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr hwnd);

    public static uint ProcessIdOfWindow(long hwnd) {
        uint pid;
        GetWindowThreadProcessId(new IntPtr(hwnd), out pid);
        return pid;
    }

    // Visible dialog boxes (#32770) of the given processes with their texts.
    public static List<string> Dialogs(uint[] pids) {
        var wanted = new HashSet<uint>(pids);
        var result = new List<string>();
        EnumWindows((hwnd, l) => {
            uint pid;
            GetWindowThreadProcessId(hwnd, out pid);
            var cls = new StringBuilder(256);
            GetClassName(hwnd, cls, cls.Capacity);
            if (!wanted.Contains(pid) || cls.ToString() != "#32770" || !IsWindowVisible(hwnd)) return true;
            var text = new StringBuilder();
            var title = new StringBuilder(512);
            GetWindowText(hwnd, title, title.Capacity);
            text.Append("[" + title + "]");
            EnumChildWindows(hwnd, (child, l2) => {
                var t = new StringBuilder(1024);
                GetWindowText(child, t, t.Capacity);
                if (t.Length > 0) text.Append(" " + t);
                return true;
            }, IntPtr.Zero);
            result.Add(text.ToString());
            return true;
        }, IntPtr.Zero);
        return result;
    }
}
'@

$exportBlock = {
    param([string]$CopyPath, [string]$RawDir)
    $ErrorActionPreference = 'Stop'
    $excel = $null
    $workbook = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        Set-Content -LiteralPath (Join-Path $RawDir 'excel.hwnd') -Value $excel.Hwnd
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.EnableEvents = $false
        $excel.AutomationSecurity = 3   # msoAutomationSecurityForceDisable
        $workbook = $excel.Workbooks.Open($CopyPath, 0, $true)

        $sheetNames = @{}
        foreach ($sheet in $workbook.Worksheets) { $sheetNames[$sheet.CodeName] = $sheet.Name }

        foreach ($component in $workbook.VBProject.VBComponents) {
            $kind = switch ($component.Type) { 1 { 'module' } 2 { 'class' } 3 { 'form' } 100 { 'document' } }
            if (-not $kind) { continue }
            $lines = $component.CodeModule.CountOfLines
            if ($kind -eq 'document' -and $lines -eq 0) { continue }

            $target = $null
            if ($kind -eq 'document') {
                if ($component.Name -eq $workbook.CodeName) { $target = 'ThisWorkbook' }
                else { $target = $sheetNames[$component.Name] }
            }
            $extension = @{ module = 'bas'; class = 'cls'; form = 'frm'; document = 'cls' }[$kind]
            $file = Join-Path $RawDir "$($component.Name).$extension"
            $component.Export($file)
            [pscustomobject]@{ Name = $component.Name; Kind = $kind; File = $file; Target = $target }
        }
    }
    finally {
        if ($workbook) { $workbook.Close($false) }
        if ($excel) { $excel.Quit() }
    }
}

$workDir = Join-Path ([IO.Path]::GetTempPath()) ('nachhilfe-vba-export-' + [guid]::NewGuid().ToString('N'))
$rawDir = Join-Path $workDir 'raw'
New-Item -ItemType Directory -Path $rawDir | Out-Null
$job = $null
$excelPid = 0

try {
    $copyPath = Join-Path $workDir (Split-Path -Leaf $WorkbookPath)
    Copy-Item -LiteralPath $WorkbookPath -Destination $copyPath
    Unblock-File -LiteralPath $copyPath

    $job = Start-Job -ScriptBlock $exportBlock -ArgumentList $copyPath, $rawDir
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    $hwndFile = Join-Path $rawDir 'excel.hwnd'
    $failure = $null

    while ($job.State -eq 'Running') {
        if (-not $excelPid -and (Test-Path -LiteralPath $hwndFile)) {
            $hwnd = (Get-Content -LiteralPath $hwndFile -ErrorAction SilentlyContinue | Select-Object -First 1)
            if ($hwnd) { $excelPid = [ExportWatch]::ProcessIdOfWindow([long]$hwnd) }
        }
        if ($excelPid) {
            $dialogs = [ExportWatch]::Dialogs([uint32[]]@($excelPid))
            if ($dialogs.Count) { $failure = 'Excel shows a dialog: ' + ($dialogs -join ' | '); break }
        }
        if ((Get-Date) -gt $deadline) { $failure = "Excel did not finish within $TimeoutSec s."; break }
        Start-Sleep -Milliseconds 500
    }

    if ($failure) {
        Stop-Job $job
        if ($excelPid) { Stop-Process -Id $excelPid -Force -ErrorAction SilentlyContinue }
        throw $failure
    }

    $components = @(Receive-Job $job -Wait -ErrorAction Stop)
    if ($job.State -ne 'Completed') { throw "Export job ended in state $($job.State)." }
    if ($excelPid) {
        $process = Get-Process -Id $excelPid -ErrorAction SilentlyContinue
        if ($process -and -not $process.WaitForExit(15000)) { Stop-Process -Id $excelPid -Force }
    }
    if (-not ($components | Where-Object Kind -eq 'module')) { throw 'The export contains no standard modules.' }

    $strictUtf8 = New-Object System.Text.UTF8Encoding($false, $true)
    $ansi = [Text.Encoding]::GetEncoding([Globalization.CultureInfo]::CurrentCulture.TextInfo.ANSICodePage)
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)

    function Read-ExportedLines([string]$Path) {
        $bytes = [IO.File]::ReadAllBytes($Path)
        try { $text = $strictUtf8.GetString($bytes) } catch { $text = $ansi.GetString($bytes) }
        return , ($text -replace '\r?\n\z', '' -split '\r?\n')
    }

    function Remove-ExportHeader([string[]]$Lines) {
        $i = 0
        while ($i -lt $Lines.Count -and $Lines[$i] -match '^(VERSION |BEGIN$|END$|\s+\w+\s*=|Attribute VB_)') { $i++ }
        if ($i -ge $Lines.Count) { return @() }
        return $Lines[$i..($Lines.Count - 1)]
    }

    function Write-SourceFile([string]$Path, [string[]]$Lines, [string]$NewLine = "`n") {
        $directory = Split-Path -Parent $Path
        if (-not (Test-Path -LiteralPath $directory)) { New-Item -ItemType Directory -Path $directory | Out-Null }
        [IO.File]::WriteAllText($Path, (($Lines -join $NewLine) + $NewLine), $utf8NoBom)
    }

    $targets = @{
        modules    = @('*.bas')
        classes    = @('*.cls')
        forms      = @('*.frm', '*.frx')
        worksheets = @('*.vba')
        workbook   = @('*.vba')
    }
    foreach ($folder in $targets.Keys) {
        $path = Join-Path $vbaRoot $folder
        if (Test-Path -LiteralPath $path) {
            foreach ($pattern in $targets[$folder]) {
                Get-ChildItem -LiteralPath $path -Filter $pattern -File | Remove-Item
            }
        }
    }

    foreach ($component in $components) {
        $lines = Read-ExportedLines $component.File
        switch ($component.Kind) {
            'module' { Write-SourceFile (Join-Path $vbaRoot "modules\$($component.Name).bas") $lines }
            'class' { Write-SourceFile (Join-Path $vbaRoot "classes\$($component.Name).cls") $lines }
            'form' {
                Write-SourceFile (Join-Path $vbaRoot "forms\$($component.Name).frm") $lines "`r`n"
                Copy-Item -LiteralPath ([IO.Path]::ChangeExtension($component.File, '.frx')) `
                    -Destination (Join-Path $vbaRoot "forms\$($component.Name).frx")
            }
            'document' {
                if (-not $component.Target) {
                    Write-Warning "Skipped document module $($component.Name): no worksheet uses it."
                }
                elseif ($component.Target -eq 'ThisWorkbook') {
                    Write-SourceFile (Join-Path $vbaRoot 'workbook\ThisWorkbook.vba') (Remove-ExportHeader $lines)
                }
                else {
                    Write-SourceFile (Join-Path $vbaRoot "worksheets\$($component.Target).vba") (Remove-ExportHeader $lines)
                }
            }
        }
    }

    $summary = $components | Group-Object Kind | ForEach-Object { "$($_.Count) $($_.Name)" }
    Write-Output "Exported from $(Split-Path -Leaf $WorkbookPath): $($summary -join ', ')."
}
finally {
    if ($job) { Remove-Job $job -Force -ErrorAction SilentlyContinue }
    Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
}
