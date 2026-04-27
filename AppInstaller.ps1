#Requires -Version 5.1
# ============================================================
#  Windows App Installer  —  AppInstaller.ps1
# ============================================================
#  Run from any PowerShell terminal:
#    powershell -ExecutionPolicy Bypass -File AppInstaller.ps1
#
#  Requirements: Windows 10/11 with winget installed.
#  Verify winget IDs anytime with:  winget search <AppName>
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ── APP DATA ──────────────────────────────────────────────────────────────────
# T  = type: winget | store | choco | ctw | url
# Id = winget package ID  |  MS Store product ID  |  URL
# S  = selected (runtime state)

$Script:Apps = @(
    @{ Name="AltSnap";                  T="winget"; Id="AltSnap.AltSnap";                  D="Move/resize windows with Alt+drag";           S=$false }
    @{ Name="Monitorian";               T="store";  Id="9NW33J738BL0";                       D="Monitor brightness controller";               S=$false }
    @{ Name="PowerToys";                T="winget"; Id="Microsoft.PowerToys";                D="Microsoft productivity utilities suite";       S=$false }
    @{ Name="Battery Mode";             T="url";    Id="https://github.com/tarcode-apps/BatteryMode/releases/latest"; D="Laptop battery plan switcher (opens browser)";  S=$false }
    @{ Name="Nilesoft Shell";           T="winget"; Id="Nilesoft.Shell";                     D="Custom right-click context menu";             S=$false }
    @{ Name="EarTrumpet";               T="winget"; Id="File-New-Project.EarTrumpet";                    D="Per-app volume control";                      S=$false }
    @{ Name="Notepads App";             T="winget"; Id="JackieLiu.NotepadsApp";                 D="Modern multi-tab notepad replacement";         S=$false }
    @{ Name="ShareX";                   T="winget"; Id="ShareX.ShareX";                      D="Screenshot, recording & file sharing";         S=$false }
    @{ Name="IconViewer";               T="url";    Id="https://www.botproductions.com/iconview/iconview.html"; D="Browse icons in EXEs/DLLs (opens browser)"; S=$false }
    @{ Name="ProtonVPN";                T="winget"; Id="Proton.ProtonVPN";        D="Secure & private VPN client";                 S=$false }
    @{ Name="Playnite";                 T="winget"; Id="Playnite.Playnite";                   D="Unified game library manager";                S=$false }
    @{ Name="Flow Launcher";            T="winget"; Id="Flow-Launcher.Flow-Launcher";         D="Keyboard launcher & quick search";            S=$false }
    @{ Name="Everything";               T="winget"; Id="voidtools.Everything";                D="Instant file & folder search";                S=$false }
    @{ Name="Clock";                    T="store";  Id="9WZDNCRFJ3PR";                       D="Windows Clock app";                           S=$false }
    @{ Name="OBS Studio";               T="winget"; Id="OBSProject.OBSStudio";                D="Screen recording & live streaming";            S=$false }
    @{ Name="QuickLook";                T="winget"; Id="QL-Win.QuickLook";                    D="Spacebar file preview in Explorer";            S=$false }
    @{ Name="BeWidgets";                T="store";  Id="9NQ07FG50H2Q";                       D="Customizable desktop widgets";                S=$false }
    @{ Name="Scrcpy";                   T="winget"; Id="Genymobile.scrcpy";                   D="Android screen mirror & remote control";       S=$false }
    @{ Name="Bulk Crap Uninstaller";    T="winget"; Id="Klocman.BulkCrapUninstaller";         D="Mass-uninstall programs & leftovers";          S=$false }
    @{ Name="FreeFileSync";             T="url";    Id="https://freefilesync.org/download.php";         D="File backup & folder sync (opens browser)";    S=$false }
    @{ Name="Wintoys";                  T="store";  Id="9P8LTPGCBZXD";                       D="Windows tweaks & cleanup tool";               S=$false }
    @{ Name="Chocolatey";               T="choco";  Id="";                                   D="Windows CLI package manager";                 S=$false }
    @{ Name="FxSound";                  T="winget"; Id="FxSound.FxSound";                    D="Real-time audio enhancement";                 S=$false }
    @{ Name="VLC";                      T="winget"; Id="VideoLAN.VLC";                        D="Universal media player";                      S=$false }
    @{ Name="Chris Titus Windows Tool"; T="ctw";    Id="";                                   D="Windows debloat & optimization (new window)"; S=$false }
    @{ Name="AutoHotkey";               T="winget"; Id="AutoHotkey.AutoHotkey";               D="Scripting & keyboard automation";             S=$false }
    @{ Name="ZoomIt";                   T="winget"; Id="Microsoft.Sysinternals.ZoomIt";       D="Screen zoom, draw & presentation timer";       S=$false }
    @{ Name="Rainmeter";                T="winget"; Id="Rainmeter.Rainmeter";                 D="Desktop skins, widgets & system gauges";       S=$false }
    @{ Name="HWiNFO";                   T="winget"; Id="REALiX.HWiNFO";                      D="Detailed hardware monitoring";                S=$false }
    @{ Name="Fan Control";              T="winget"; Id="Rem0o.FanControl";                    D="Advanced CPU/GPU fan curve control";           S=$false }
    @{ Name="Mica For Everyone";        T="winget"; Id="MicaForEveryone.MicaForEveryone";     D="Mica transparency for any window";             S=$false }
    @{ Name="MSI Afterburner";          T="winget"; Id="Guru3D.Afterburner";                  D="GPU overclocking & monitoring";               S=$false }
    @{ Name="SuperF4";                  T="winget"; Id="StefanSundin.Superf4";                D="Force-kill any frozen application";           S=$false }
    @{ Name="Keyviz";                   T="winget"; Id="mulaRahul.Keyviz";                   D="Real-time keystroke visualizer on screen";     S=$false }
)

# Thread-safe message queue (background runspace → UI timer)
$Script:Q = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()

# ── PALETTE ───────────────────────────────────────────────────────────────────
$C = @{
    Bg       = [Drawing.Color]::FromArgb(15, 15, 15)
    BgPanel  = [Drawing.Color]::FromArgb(22, 22, 22)
    BgRow1   = [Drawing.Color]::FromArgb(26, 26, 26)
    BgRow2   = [Drawing.Color]::FromArgb(32, 32, 32)
    BgInput  = [Drawing.Color]::FromArgb(38, 38, 38)
    BgBtn    = [Drawing.Color]::FromArgb(44, 44, 44)
    Accent   = [Drawing.Color]::FromArgb(0, 112, 199)
    AccentHv = [Drawing.Color]::FromArgb(0, 136, 232)
    White    = [Drawing.Color]::FromArgb(228, 228, 228)
    Muted    = [Drawing.Color]::FromArgb(105, 105, 105)
    Green    = [Drawing.Color]::FromArgb(72, 199, 116)
    Yellow   = [Drawing.Color]::FromArgb(220, 185, 50)
    Red      = [Drawing.Color]::FromArgb(220, 70, 70)
    Blue     = [Drawing.Color]::FromArgb(99, 179, 237)
    Border   = [Drawing.Color]::FromArgb(48, 48, 48)
}

# ── FORM ──────────────────────────────────────────────────────────────────────
$form = New-Object Windows.Forms.Form
$form.Text            = "Windows App Installer"
$form.ClientSize      = New-Object Drawing.Size(760, 790)
$form.StartPosition   = "CenterScreen"
$form.BackColor       = $C.Bg
$form.ForeColor       = $C.White
$form.Font            = New-Object Drawing.Font("Segoe UI", 10)
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox     = $false

# ── HEADER ────────────────────────────────────────────────────────────────────
$pnlHead          = New-Object Windows.Forms.Panel
$pnlHead.Location = New-Object Drawing.Point(0, 0)
$pnlHead.Size     = New-Object Drawing.Size(760, 70)
$pnlHead.BackColor= $C.BgPanel
$form.Controls.Add($pnlHead)

$lbTitle          = New-Object Windows.Forms.Label
$lbTitle.Text     = "  ⚡  Windows App Installer"
$lbTitle.Font     = New-Object Drawing.Font("Segoe UI", 15, [Drawing.FontStyle]::Bold)
$lbTitle.ForeColor= $C.Blue
$lbTitle.AutoSize = $true
$lbTitle.Location = New-Object Drawing.Point(10, 10)
$pnlHead.Controls.Add($lbTitle)

$lbSub            = New-Object Windows.Forms.Label
$lbSub.Text       = "  Select apps below, then click Install. Powered by winget, MS Store & custom scripts."
$lbSub.ForeColor  = $C.Muted
$lbSub.Font       = New-Object Drawing.Font("Segoe UI", 9)
$lbSub.AutoSize   = $true
$lbSub.Location   = New-Object Drawing.Point(10, 46)
$pnlHead.Controls.Add($lbSub)

# ── TOOLBAR ───────────────────────────────────────────────────────────────────
$pnlBar          = New-Object Windows.Forms.Panel
$pnlBar.Location = New-Object Drawing.Point(0, 70)
$pnlBar.Size     = New-Object Drawing.Size(760, 50)
$pnlBar.BackColor= [Drawing.Color]::FromArgb(20, 20, 20)
$form.Controls.Add($pnlBar)

function Make-Btn($txt, $x, $w=96, $h=30) {
    $b = New-Object Windows.Forms.Button
    $b.Text      = $txt
    $b.Location  = New-Object Drawing.Point($x, 10)
    $b.Size      = New-Object Drawing.Size($w, $h)
    $b.BackColor = $C.BgBtn
    $b.ForeColor = $C.White
    $b.FlatStyle = "Flat"
    $b.FlatAppearance.BorderColor         = $C.Border
    $b.FlatAppearance.MouseOverBackColor  = [Drawing.Color]::FromArgb(60, 60, 60)
    $b.Cursor    = "Hand"
    return $b
}

$btnAll  = Make-Btn "☑  All"  12
$btnNone = Make-Btn "☐  None" 114
$pnlBar.Controls.AddRange(@($btnAll, $btnNone))

$lbSrchLbl          = New-Object Windows.Forms.Label
$lbSrchLbl.Text     = "Search:"
$lbSrchLbl.ForeColor= $C.Muted
$lbSrchLbl.AutoSize = $true
$lbSrchLbl.Location = New-Object Drawing.Point(222, 16)
$pnlBar.Controls.Add($lbSrchLbl)

$txtSearch             = New-Object Windows.Forms.TextBox
$txtSearch.Location    = New-Object Drawing.Point(276, 13)
$txtSearch.Size        = New-Object Drawing.Size(210, 26)
$txtSearch.BackColor   = $C.BgInput
$txtSearch.ForeColor   = $C.White
$txtSearch.BorderStyle = "FixedSingle"
$pnlBar.Controls.Add($txtSearch)

$lbCount           = New-Object Windows.Forms.Label
$lbCount.ForeColor = $C.Muted
$lbCount.Font      = New-Object Drawing.Font("Segoe UI", 9)
$lbCount.AutoSize  = $true
$lbCount.Location  = New-Object Drawing.Point(500, 16)
$pnlBar.Controls.Add($lbCount)

# ── APP LIST ──────────────────────────────────────────────────────────────────
$pnlList           = New-Object Windows.Forms.Panel
$pnlList.Location  = New-Object Drawing.Point(10, 128)
$pnlList.Size      = New-Object Drawing.Size(740, 450)
$pnlList.BackColor = $C.Bg
$pnlList.AutoScroll= $true
$pnlList.BorderStyle = "None"
$form.Controls.Add($pnlList)

function Update-Count {
    $n = ($Script:Apps | Where-Object { $_.S }).Count
    $lbCount.Text = "$n / $($Script:Apps.Count) selected"
}

function Build-List {
    param([string]$filter = "")
    $pnlList.SuspendLayout()
    $pnlList.Controls.Clear()
    $y = 1; $alt = $false

    foreach ($app in $Script:Apps) {
        if ($filter -and
            $app.Name -notlike "*$filter*" -and
            $app.D    -notlike "*$filter*") { continue }

        $alt = -not $alt
        $rowBg = if ($alt) { $C.BgRow1 } else { $C.BgRow2 }

        $row           = New-Object Windows.Forms.Panel
        $row.Location  = New-Object Drawing.Point(1, $y)
        $row.Size      = New-Object Drawing.Size(718, 46)
        $row.BackColor = $rowBg

        # Checkbox
        $cb          = New-Object Windows.Forms.CheckBox
        $cb.Checked  = $app.S
        $cb.Location = New-Object Drawing.Point(12, 14)
        $cb.Size     = New-Object Drawing.Size(18, 18)
        $cb.BackColor= [Drawing.Color]::Transparent

        # App name
        $lbN           = New-Object Windows.Forms.Label
        $lbN.Text      = $app.Name
        $lbN.Location  = New-Object Drawing.Point(40, 6)
        $lbN.Size      = New-Object Drawing.Size(295, 18)
        $lbN.Font      = New-Object Drawing.Font("Segoe UI", 10, [Drawing.FontStyle]::Bold)
        $lbN.ForeColor = $C.White
        $lbN.BackColor = [Drawing.Color]::Transparent

        # Description
        $lbD           = New-Object Windows.Forms.Label
        $lbD.Text      = $app.D
        $lbD.Location  = New-Object Drawing.Point(40, 26)
        $lbD.Size      = New-Object Drawing.Size(460, 16)
        $lbD.Font      = New-Object Drawing.Font("Segoe UI", 8.5)
        $lbD.ForeColor = $C.Muted
        $lbD.BackColor = [Drawing.Color]::Transparent

        # Type badge
        $badge          = New-Object Windows.Forms.Label
        $badge.Text     = switch ($app.T) {
            "winget" { "WINGET"   }
            "store"  { "MS STORE" }
            "choco"  { "CHOCO"    }
            "ctw"    { "SCRIPT"   }
            "url"    { "BROWSER"  }
        }
        $badge.BackColor = switch ($app.T) {
            "winget" { [Drawing.Color]::FromArgb(0,  71, 133) }
            "store"  { [Drawing.Color]::FromArgb(0,  90,  55) }
            "choco"  { [Drawing.Color]::FromArgb(105, 50,  0) }
            "ctw"    { [Drawing.Color]::FromArgb(90,  20,  90) }
            "url"    { [Drawing.Color]::FromArgb(70,  70,  10) }
        }
        $badge.ForeColor  = [Drawing.Color]::White
        $badge.Font       = New-Object Drawing.Font("Segoe UI", 7, [Drawing.FontStyle]::Bold)
        $badge.TextAlign  = [Drawing.ContentAlignment]::MiddleCenter
        $badge.Location   = New-Object Drawing.Point(630, 14)
        $badge.Size       = New-Object Drawing.Size(74, 18)

        $row.Controls.AddRange(@($cb, $lbN, $lbD, $badge))

        # Wire events — GetNewClosure() fixes the closure-in-loop problem
        $capturedApp = $app
        $capturedCb  = $cb

        $cb.Add_CheckedChanged(
            { $capturedApp.S = $this.Checked; Update-Count }.GetNewClosure()
        )
        foreach ($ctl in @($row, $lbN, $lbD, $badge)) {
            $ctl.Cursor = "Hand"
            $ctl.Add_Click(
                { $capturedCb.Checked = -not $capturedCb.Checked }.GetNewClosure()
            )
        }

        $pnlList.Controls.Add($row)
        $y += 47
    }

    $pnlList.AutoScrollMinSize = New-Object Drawing.Size(0, ($y + 4))
    $pnlList.ResumeLayout()
    Update-Count
}

Build-List

$btnAll.Add_Click({
    foreach ($a in $Script:Apps) { $a.S = $true }
    Build-List -filter $txtSearch.Text
})
$btnNone.Add_Click({
    foreach ($a in $Script:Apps) { $a.S = $false }
    Build-List -filter $txtSearch.Text
})
$txtSearch.Add_TextChanged({ Build-List -filter $txtSearch.Text })

# ── SEPARATOR ─────────────────────────────────────────────────────────────────
$sep           = New-Object Windows.Forms.Panel
$sep.Location  = New-Object Drawing.Point(0, 584)
$sep.Size      = New-Object Drawing.Size(760, 1)
$sep.BackColor = $C.Border
$form.Controls.Add($sep)

# ── LOG ───────────────────────────────────────────────────────────────────────
$lbLog          = New-Object Windows.Forms.Label
$lbLog.Text     = "INSTALL LOG"
$lbLog.ForeColor= $C.Muted
$lbLog.Font     = New-Object Drawing.Font("Segoe UI", 7.5, [Drawing.FontStyle]::Bold)
$lbLog.AutoSize = $true
$lbLog.Location = New-Object Drawing.Point(12, 591)
$form.Controls.Add($lbLog)

$rtLog             = New-Object Windows.Forms.RichTextBox
$rtLog.Location    = New-Object Drawing.Point(10, 611)
$rtLog.Size        = New-Object Drawing.Size(740, 120)
$rtLog.BackColor   = [Drawing.Color]::FromArgb(10, 10, 10)
$rtLog.ForeColor   = $C.Green
$rtLog.Font        = New-Object Drawing.Font("Consolas", 8.5)
$rtLog.ReadOnly    = $true
$rtLog.BorderStyle = "None"
$rtLog.ScrollBars  = "Vertical"
$form.Controls.Add($rtLog)

function Log {
    param($msg, $col = $null)
    if (-not $col) { $col = $C.Muted }
    $rtLog.SelectionStart  = $rtLog.TextLength
    $rtLog.SelectionLength = 0
    $rtLog.SelectionColor  = $col
    $rtLog.AppendText("$msg`n")
    $rtLog.ScrollToCaret()
}

# ── BOTTOM BAR ────────────────────────────────────────────────────────────────
$pnlBot           = New-Object Windows.Forms.Panel
$pnlBot.Location  = New-Object Drawing.Point(0, 740)
$pnlBot.Size      = New-Object Drawing.Size(760, 50)
$pnlBot.BackColor = $C.BgPanel
$form.Controls.Add($pnlBot)

$pbar          = New-Object Windows.Forms.ProgressBar
$pbar.Location = New-Object Drawing.Point(10, 13)
$pbar.Size     = New-Object Drawing.Size(548, 24)
$pbar.Style    = "Continuous"
$pnlBot.Controls.Add($pbar)

$btnInst = Make-Btn "  Install ▶" 570 178 34
$btnInst.Location  = New-Object Drawing.Point(570, 8)
$btnInst.BackColor = $C.Accent
$btnInst.Font      = New-Object Drawing.Font("Segoe UI", 11, [Drawing.FontStyle]::Bold)
$btnInst.FlatAppearance.BorderSize         = 0
$btnInst.FlatAppearance.MouseOverBackColor = $C.AccentHv
$pnlBot.Controls.Add($btnInst)

# ── POLL TIMER (drains bg runspace messages onto UI thread) ───────────────────
$Script:Timer          = New-Object Windows.Forms.Timer
$Script:Timer.Interval = 80
$Script:Timer.Add_Tick({
    $msg = $null
    while ($Script:Q.TryDequeue([ref]$msg)) {
        switch -Regex ($msg) {
            "^__DONE__$" {
                $Script:Timer.Stop()
                $btnInst.Enabled = $true
                $btnInst.Text    = "  Install ▶"
                $pbar.Value      = $pbar.Maximum
                Log ""
                Log "  ✔  All installations finished! Check the log above for results." $C.Green
            }
            "^__PROG:(\d+)$" {
                $v = [int]$Matches[1]
                if ($v -le $pbar.Maximum) { $pbar.Value = $v }
            }
            "^__OK:(.+)$"   { Log $Matches[1] $C.Green  }
            "^__ERR:(.+)$"  { Log $Matches[1] $C.Red    }
            "^__WARN:(.+)$" { Log $Matches[1] $C.Yellow }
            "^__HEAD:(.+)$" { Log $Matches[1] $C.Blue   }
            default         { Log $msg         $C.Muted  }
        }
    }
})

# ── INSTALL BUTTON ────────────────────────────────────────────────────────────
$btnInst.Add_Click({
    $sel = @($Script:Apps | Where-Object { $_.S })

    if ($sel.Count -eq 0) {
        [Windows.Forms.MessageBox]::Show(
            "Select at least one app to install.",
            "Nothing Selected", "OK", "Information"
        ) | Out-Null
        return
    }

    $listText = ($sel | ForEach-Object { "  • $($_.Name)" }) -join "`n"
    $res = [Windows.Forms.MessageBox]::Show(
        "Ready to install $($sel.Count) app(s):`n`n$listText`n`nContinue?",
        "Confirm Install", "YesNo", "Question"
    )
    if ($res -ne "Yes") { return }

    # Reset UI
    $btnInst.Enabled = $false
    $btnInst.Text    = "  Working..."
    $rtLog.Clear()
    $pbar.Value      = 0
    $pbar.Maximum    = $sel.Count
    $Script:Q        = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()

    # Snapshot selected app data (plain PSCustomObjects cross runspace boundaries cleanly)
    $appData = @($sel | ForEach-Object {
        [PSCustomObject]@{ Name=$_.Name; T=$_.T; Id=$_.Id }
    })

    # Create background runspace
    $rs = [RunspaceFactory]::CreateRunspace()
    $rs.ApartmentState = "STA"
    $rs.ThreadOptions  = "ReuseThread"
    $rs.Open()
    $rs.SessionStateProxy.SetVariable("apps", $appData)
    $rs.SessionStateProxy.SetVariable("q",    $Script:Q)

    $ps = [PowerShell]::Create()
    $ps.Runspace = $rs

    [void]$ps.AddScript({
        $total = $apps.Count
        $i     = 0

        foreach ($app in $apps) {
            $i++
            $q.Enqueue("__HEAD:  [$i/$total]  $($app.Name)")

            try {
                switch ($app.T) {

                    # ── winget ────────────────────────────────────────────
                    "winget" {
                        $out  = & winget install --id "$($app.Id)" -e `
                                    --accept-source-agreements `
                                    --accept-package-agreements `
                                    --silent 2>&1
                        $code = $LASTEXITCODE

                        # Show last few non-empty output lines
                        $out | Select-Object -Last 4 | ForEach-Object {
                            $ln = "$_".Trim()
                            if ($ln) { $q.Enqueue("  $ln") }
                        }

                        if     ($code -eq 0)           { $q.Enqueue("__OK:  ✔  $($app.Name) installed successfully.") }
                        elseif ($code -eq -1978335189) { $q.Enqueue("__WARN:  ⚠  $($app.Name) is already installed.") }
                        else                           { $q.Enqueue("__ERR:  ✘  $($app.Name) failed (exit code $code).") }
                    }

                    # ── MS Store via winget ───────────────────────────────
                    "store" {
                        $out  = & winget install --id "$($app.Id)" -e `
                                    --source msstore `
                                    --accept-source-agreements `
                                    --accept-package-agreements 2>&1
                        $code = $LASTEXITCODE

                        if     ($code -eq 0)           { $q.Enqueue("__OK:  ✔  $($app.Name) installed successfully.") }
                        elseif ($code -eq -1978335189) { $q.Enqueue("__WARN:  ⚠  $($app.Name) is already installed.") }
                        else                           { $q.Enqueue("__ERR:  ✘  $($app.Name) failed (exit code $code).") }
                    }

                    # ── Chocolatey bootstrap ──────────────────────────────
                    "choco" {
                        $q.Enqueue("  Downloading Chocolatey install script...")
                        [Net.ServicePointManager]::SecurityProtocol =
                            [Net.ServicePointManager]::SecurityProtocol -bor 3072
                        $script = (New-Object Net.WebClient).DownloadString(
                            'https://community.chocolatey.org/install.ps1'
                        )
                        Invoke-Expression $script 2>&1 | Out-Null

                        if (Get-Command choco -ErrorAction SilentlyContinue) {
                            $q.Enqueue("__OK:  ✔  Chocolatey installed. Restart your terminal to use 'choco'.")
                        } else {
                            $q.Enqueue("__ERR:  ✘  Chocolatey may not have installed correctly. Try running as admin.")
                        }
                    }

                    # ── Chris Titus Tool (interactive — new window) ───────
                    "ctw" {
                        $q.Enqueue("  Launching Chris Titus Tool in a new admin PowerShell window...")
                        Start-Process powershell.exe `
                            -ArgumentList '-NoProfile -Command "irm christitus.com/win | iex"' `
                            -Verb RunAs
                        $q.Enqueue("__OK:  ✔  Chris Titus Tool launched in a separate window.")
                    }

                    # ── URL — open download page ──────────────────────────
                    "url" {
                        $q.Enqueue("  Opening download page in your browser...")
                        Start-Process $app.Id
                        $q.Enqueue("__OK:  ✔  Browser opened for $($app.Name). Download manually from the page.")
                    }
                }
            } catch {
                $q.Enqueue("__ERR:  ✘  $($app.Name) threw an exception: $_")
            }

            $q.Enqueue("__PROG:$i")
            $q.Enqueue("")   # blank line between apps
        }

        $q.Enqueue("__DONE__")
    })

    $Script:Timer.Start()
    [void]$ps.BeginInvoke()
})

# ── STARTUP CHECK ─────────────────────────────────────────────────────────────
$form.Add_Shown({
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        [Windows.Forms.MessageBox]::Show(
            "winget (Windows Package Manager) was not found on this system.`n`n" +
            "Please install it from the Microsoft Store (search 'App Installer')`n" +
            "or update Windows, then re-run this script.`n`n" +
            "Note: apps of type CHOCO, SCRIPT, and BROWSER don't need winget.",
            "winget Not Found", "OK", "Warning"
        ) | Out-Null
    }
})

$form.Add_FormClosing({ $Script:Timer.Stop() })

# ── LAUNCH ────────────────────────────────────────────────────────────────────
[Windows.Forms.Application]::Run($form)
