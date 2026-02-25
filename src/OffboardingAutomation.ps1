# UTF-8 output (Dosyayi UTF-8 kaydedin)
# Save file as UTF-8
try {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $OutputEncoding = [System.Text.Encoding]::UTF8
} catch {}

# STA (WinForms stabil)
# STA for WinForms stability
if ([threading.thread]::CurrentThread.ApartmentState -ne 'STA') {
    $ps = (Get-Process -Id $PID).Path
    Start-Process $ps -ArgumentList "-NoProfile -ExecutionPolicy Bypass -STA -File `"$PSCommandPath`""
    exit
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Rounded form corners
# Form oval koseler
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Win32 {
  [DllImport("gdi32.dll", SetLastError=true)]
  public static extern IntPtr CreateRoundRectRgn(int nLeftRect, int nTopRect, int nRightRect, int nBottomRect, int nWidthEllipse, int nHeightEllipse);
}
"@

function Set-FormRoundedCorners {
    param([System.Windows.Forms.Form]$Form, [int]$Radius = 18)
    $rgn = [Win32]::CreateRoundRectRgn(0, 0, $Form.Width + 1, $Form.Height + 1, $Radius, $Radius)
    $Form.Region = [System.Drawing.Region]::FromHrgn($rgn)
}

function Set-RoundedCorners {
    param($Control, [int]$Radius)
    $Path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $Path.AddArc(0, 0, $Radius, $Radius, 180, 90)
    $Path.AddArc($Control.Width - $Radius, 0, $Radius, $Radius, 270, 90)
    $Path.AddArc($Control.Width - $Radius, $Control.Height - $Radius, $Radius, $Radius, 0, 90)
    $Path.AddArc(0, $Control.Height - $Radius, $Radius, $Radius, 90, 90)
    $Path.CloseFigure()
    $Control.Region = New-Object System.Drawing.Region($Path)
    if ($Control -is [System.Windows.Forms.Button]) {
        $Control.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $Control.FlatAppearance.BorderSize = 0
    }
}

# Ana ekran
# Main screen
$Global:Lang = "TR"
$Global:SelectedUser = $null
$Global:SelectedUserId = $null
$Global:SelectedSMTP = $null

# Dil metinleri
# Localization strings
$L = @{
    TR = @{
        AppTitle             = "Offboarding Automation v1"
        BtnModules           = "Modulleri Yukle/Guncelle"
        BtnADConnect         = "AD Baglan"
        ADReady              = "AD Hazir"
        GroupUser            = "1. Hedef Kullaniciyi Secin"
        LblSearch            = "Ara (en az 2 harf):"
        BtnLock              = "Kilitle"
        BtnUnlock            = "Ac"
        GroupAD              = "2. AD On-Prem"
        GroupEXO             = "3. Exchange Online + Purview"
        GroupGraph           = "4. Graph / Entra ID"
        BtnADOnly            = "SADECE AD"
        BtnEXOOnly           = "SADECE EXO/PURVIEW"
        BtnGraphOnly         = "SADECE GRAPH"
        BtnRunAll            = "HEPSINI CALISTIR (AD -> EXO/PURVIEW -> GRAPH)"

        ChkADDisable         = "AD hesabini devre disi birak"
        ChkADGroups          = "Tum AD gruplarindan cikar"
        ChkADPwd             = "Parolayi sifirla (random)"
        ChkADDesc            = "Description guncelle"
        ChkADGal             = "GAL temizle (Hide+Clear)"
        ChkADOrg             = "Organization temizle (Manager/Reports)"

        ChkCal               = "Takvim etkinliklerini sil (3 retry)"
        ChkShared            = "Shared Mailbox'a cevir"
        ChkPurview           = "Purview Search olustur + calistir"

        ChkRevoke            = "Oturumlari dusur (Revoke)"
        ChkGraphGroups       = "M365/Teams gruplarindan cikar"
        ChkLic               = "Lisanslari geri al"

        LogReady             = "Sistem hazir. Kullanici ara -> sec -> Kilitle. (Kilit yoksa butonlar pasif)"
        LogSelected          = "Secildi"
        LogLocked            = "Kullanici kilitlendi. (Degistirmek icin Ac)"
        LogUnlocked          = "Kilit acildi. Yeni kullanici secebilirsin."

        MsgPrivTitle         = "Guvenlik Engeli"
        MsgPrivBody          = "Secilen kullanici kritik admin gruplarinda (Domain/Enterprise/Schema Admins veya Builtin Administrators).`n`nGuvenlik nedeniyle islemler durduruldu.`n`nOnce AD'den admin yetkilerini kaldirin, sonra tekrar deneyin."

        LogPurviewTitle      = "PST export icin (UI uzerinden):"
        LogPurview1          = "1) https://purview.microsoft.com/ adresine giris yapin."
        LogPurview2          = "2) Dava dosyalari > Icerik arama (Content Search) yoluna gidin."
        LogPurview3          = "3) Arama adindan '{0}' aratin / acin."
        LogPurview4          = "4) 'Disari aktar' diyerek PST paketini olusturup indirebilirsiniz."
        LogPurviewNote       = "Not: Export islemi Purview arayuzunden yapilir."

        LogRunAllStart       = "##### FULL OFFBOARD START (AD -> EXO/PURVIEW -> GRAPH) #####"
        LogRunAllEnd         = "##### FULL OFFBOARD END #####"
    }
    EN = @{
        AppTitle             = "Offboarding Automation v1"
        BtnModules           = "Install/Update Modules"
        BtnADConnect         = "Connect AD"
        ADReady              = "AD Ready"
        GroupUser            = "1. Select Target User"
        LblSearch            = "Search (min 2 chars):"
        BtnLock              = "Lock"
        BtnUnlock            = "Unlock"
        GroupAD              = "2. AD On-Prem"
        GroupEXO             = "3. Exchange Online + Purview"
        GroupGraph           = "4. Graph / Entra ID"
        BtnADOnly            = "AD ONLY"
        BtnEXOOnly           = "EXO/PURVIEW ONLY"
        BtnGraphOnly         = "GRAPH ONLY"
        BtnRunAll            = "RUN ALL (AD -> EXO/PURVIEW -> GRAPH)"

        ChkADDisable         = "Disable AD account"
        ChkADGroups          = "Remove from all AD groups"
        ChkADPwd             = "Reset password (random)"
        ChkADDesc            = "Update description"
        ChkADGal             = "Hide from GAL (Hide+Clear)"
        ChkADOrg             = "Org cleanup (Manager/Reports)"

        ChkCal               = "Remove calendar events (3 retries)"
        ChkShared            = "Convert to shared mailbox"
        ChkPurview           = "Create + start Purview search"

        ChkRevoke            = "Revoke sign-in sessions"
        ChkGraphGroups       = "Remove from M365/Teams groups"
        ChkLic               = "Remove licenses"

        LogReady             = "Ready. Search -> select -> Lock. (Buttons disabled unless locked)"
        LogSelected          = "Selected"
        LogLocked            = "User locked. (Unlock to change)"
        LogUnlocked          = "Unlocked. You can select a new user."

        MsgPrivTitle         = "Security Block"
        MsgPrivBody          = "Selected user is in privileged admin groups (Domain/Enterprise/Schema Admins or Builtin Administrators).`n`nOperations are blocked for safety.`n`nRemove admin rights first, then try again."

        LogPurviewTitle      = "For PST export (via UI):"
        LogPurview1          = "1) Sign in to https://purview.microsoft.com/"
        LogPurview2          = "2) Go to Cases > Content search."
        LogPurview3          = "3) Open the search named '{0}'."
        LogPurview4          = "4) Click 'Export' to generate and download the PST package."
        LogPurviewNote       = "Note: Export is done via the Purview UI."

        LogRunAllStart       = "##### FULL OFFBOARD START (AD -> EXO/PURVIEW -> GRAPH) #####"
        LogRunAllEnd         = "##### FULL OFFBOARD END #####"
    }
}

function T { param([string]$k) return $L[$Global:Lang][$k] }

# Log
$Form = New-Object System.Windows.Forms.Form
$Form.Text = (T "AppTitle")
$Form.Size = New-Object System.Drawing.Size(1000, 650)
$Form.StartPosition = "CenterScreen"
$Form.BackColor = [System.Drawing.Color]::WhiteSmoke
$Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
$Form.MaximizeBox = $false
$Form.ShowIcon = $false

$topBar = New-Object System.Windows.Forms.Panel
$topBar.Location = New-Object System.Drawing.Point(0,0)
$topBar.Size = New-Object System.Drawing.Size(1000,40)
$topBar.BackColor = [System.Drawing.Color]::Gainsboro
$Form.Controls.Add($topBar)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = $Form.Text
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$lblTitle.AutoSize = $true
$lblTitle.Location = New-Object System.Drawing.Point(12,11)
$topBar.Controls.Add($lblTitle)

# Dil secimi (sag ust)
# Language switch (top right)
$cmbLang = New-Object System.Windows.Forms.ComboBox
$cmbLang.DropDownStyle = 'DropDownList'
$cmbLang.Items.AddRange(@("TR","EN"))
$cmbLang.SelectedIndex = 0
$cmbLang.Location = New-Object System.Drawing.Point(880,9)
$cmbLang.Size = New-Object System.Drawing.Size(60,22)
$topBar.Controls.Add($cmbLang)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "X"
$btnClose.Location = New-Object System.Drawing.Point(960,6)
$btnClose.Size = New-Object System.Drawing.Size(30,28)
$btnClose.BackColor = [System.Drawing.Color]::LightCoral
Set-RoundedCorners $btnClose 10
$topBar.Controls.Add($btnClose)
$btnClose.Add_Click({ $Form.Close() })

# Drag
$script:isDragging = $false
$script:dragOffset = New-Object System.Drawing.Point(0,0)
$topBar.Add_MouseDown({ param($s,$e) if ($e.Button -eq 'Left') { $script:isDragging=$true; $script:dragOffset=$e.Location } })
$topBar.Add_MouseMove({
    param($s,$e)
    if ($script:isDragging) {
        $p = [System.Windows.Forms.Control]::MousePosition
        $Form.Location = New-Object System.Drawing.Point(($p.X - $script:dragOffset.X), ($p.Y - $script:dragOffset.Y))
    }
})
$topBar.Add_MouseUp({ param($s,$e) $script:isDragging=$false })

$rtbLog = New-Object System.Windows.Forms.RichTextBox
$rtbLog.Location = New-Object System.Drawing.Point(20, 380)
$rtbLog.Size = New-Object System.Drawing.Size(940, 200)
$rtbLog.BackColor = [System.Drawing.Color]::Black
$rtbLog.ForeColor = [System.Drawing.Color]::LightGray
$rtbLog.ReadOnly = $true
$rtbLog.Font = New-Object System.Drawing.Font("Consolas", 10)
Set-RoundedCorners $rtbLog 15
$Form.Controls.Add($rtbLog)

function Write-Log {
    param([string]$Message, [string]$Color = "LightGray")
    $rtbLog.SelectionStart = $rtbLog.TextLength
    $rtbLog.SelectionLength = 0
    $rtbLog.SelectionColor = [System.Drawing.Color]::$Color
    $rtbLog.AppendText("$Message`n")
    $rtbLog.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

# Sag ust butonlar
$btnModules = New-Object System.Windows.Forms.Button
$btnModules.Text = (T "BtnModules")
$btnModules.Location = New-Object System.Drawing.Point(760, 55)
$btnModules.Size = New-Object System.Drawing.Size(200, 30)
$btnModules.BackColor = [System.Drawing.Color]::LightGray
Set-RoundedCorners $btnModules 10
$Form.Controls.Add($btnModules)

$btnADConnect = New-Object System.Windows.Forms.Button
$btnADConnect.Text = (T "BtnADConnect")
$btnADConnect.Location = New-Object System.Drawing.Point(760, 95)
$btnADConnect.Size = New-Object System.Drawing.Size(200, 30)
$btnADConnect.BackColor = [System.Drawing.Color]::LightCoral
Set-RoundedCorners $btnADConnect 10
$Form.Controls.Add($btnADConnect)

# Kullanici secimi
$grpUser = New-Object System.Windows.Forms.GroupBox
$grpUser.Text = (T "GroupUser")
$grpUser.Location = New-Object System.Drawing.Point(20, 55)
$grpUser.Size = New-Object System.Drawing.Size(720, 120)
$Form.Controls.Add($grpUser)

$lblSearch = New-Object System.Windows.Forms.Label
$lblSearch.Text = (T "LblSearch")
$lblSearch.Location = New-Object System.Drawing.Point(20, 25)
$lblSearch.AutoSize = $true
$grpUser.Controls.Add($lblSearch)

$txtSearch = New-Object System.Windows.Forms.TextBox
$txtSearch.Location = New-Object System.Drawing.Point(20, 45)
$txtSearch.Size = New-Object System.Drawing.Size(250, 22)
$grpUser.Controls.Add($txtSearch)

$txtSMTP = New-Object System.Windows.Forms.TextBox
$txtSMTP.Location = New-Object System.Drawing.Point(290, 45)
$txtSMTP.Size = New-Object System.Drawing.Size(350, 22)
$txtSMTP.ReadOnly = $true
$grpUser.Controls.Add($txtSMTP)

$btnLock = New-Object System.Windows.Forms.Button
$btnLock.Text = (T "BtnLock")
$btnLock.Location = New-Object System.Drawing.Point(650, 42)
$btnLock.Size = New-Object System.Drawing.Size(60, 28)
$btnLock.BackColor = [System.Drawing.Color]::Khaki
Set-RoundedCorners $btnLock 10
$grpUser.Controls.Add($btnLock)

$lstUsers = New-Object System.Windows.Forms.ListBox
$lstUsers.Location = New-Object System.Drawing.Point(20, 70)
$lstUsers.Size = New-Object System.Drawing.Size(690, 40)
$grpUser.Controls.Add($lstUsers)

# AD On-Prem
$grpAD = New-Object System.Windows.Forms.GroupBox
$grpAD.Text = (T "GroupAD")
$grpAD.Location = New-Object System.Drawing.Point(20, 190)
$grpAD.Size = New-Object System.Drawing.Size(300, 180)
$Form.Controls.Add($grpAD)

$chkADDisable = New-Object System.Windows.Forms.CheckBox
$chkADDisable.Text = (T "ChkADDisable")
$chkADDisable.Location = New-Object System.Drawing.Point(15, 25)
$chkADDisable.Width = 280
$grpAD.Controls.Add($chkADDisable)

$chkADGroups = New-Object System.Windows.Forms.CheckBox
$chkADGroups.Text = (T "ChkADGroups")
$chkADGroups.Location = New-Object System.Drawing.Point(15, 50)
$chkADGroups.Width = 280
$grpAD.Controls.Add($chkADGroups)

$chkADPwd = New-Object System.Windows.Forms.CheckBox
$chkADPwd.Text = (T "ChkADPwd")
$chkADPwd.Location = New-Object System.Drawing.Point(15, 75)
$chkADPwd.Width = 280
$grpAD.Controls.Add($chkADPwd)

$chkADDesc = New-Object System.Windows.Forms.CheckBox
$chkADDesc.Text = (T "ChkADDesc")
$chkADDesc.Location = New-Object System.Drawing.Point(15, 100)
$chkADDesc.Width = 280
$grpAD.Controls.Add($chkADDesc)

$chkADGal = New-Object System.Windows.Forms.CheckBox
$chkADGal.Text = (T "ChkADGal")
$chkADGal.Location = New-Object System.Drawing.Point(15, 125)
$chkADGal.Width = 280
$grpAD.Controls.Add($chkADGal)

$chkADOrg = New-Object System.Windows.Forms.CheckBox
$chkADOrg.Text = (T "ChkADOrg")
$chkADOrg.Location = New-Object System.Drawing.Point(15, 150)
$chkADOrg.Width = 280
$chkADOrg.Checked = $true
$grpAD.Controls.Add($chkADOrg)

# EXO + Purview
$grpEXO = New-Object System.Windows.Forms.GroupBox
$grpEXO.Text = (T "GroupEXO")
$grpEXO.Location = New-Object System.Drawing.Point(340, 190)
$grpEXO.Size = New-Object System.Drawing.Size(300, 180)
$Form.Controls.Add($grpEXO)

$chkCal = New-Object System.Windows.Forms.CheckBox
$chkCal.Text = (T "ChkCal")
$chkCal.Location = New-Object System.Drawing.Point(15, 30)
$chkCal.Width = 280
$chkCal.Checked = $true
$grpEXO.Controls.Add($chkCal)

$chkShared = New-Object System.Windows.Forms.CheckBox
$chkShared.Text = (T "ChkShared")
$chkShared.Location = New-Object System.Drawing.Point(15, 60)
$chkShared.Width = 280
$chkShared.Checked = $true
$grpEXO.Controls.Add($chkShared)

$chkPurview = New-Object System.Windows.Forms.CheckBox
$chkPurview.Text = (T "ChkPurview")
$chkPurview.Location = New-Object System.Drawing.Point(15, 90)
$chkPurview.Width = 280
$chkPurview.Checked = $true
$grpEXO.Controls.Add($chkPurview)

# Graph / Entra
$grpGraph = New-Object System.Windows.Forms.GroupBox
$grpGraph.Text = (T "GroupGraph")
$grpGraph.Location = New-Object System.Drawing.Point(660, 190)
$grpGraph.Size = New-Object System.Drawing.Size(300, 180)
$Form.Controls.Add($grpGraph)

$chkRevoke = New-Object System.Windows.Forms.CheckBox
$chkRevoke.Text = (T "ChkRevoke")
$chkRevoke.Location = New-Object System.Drawing.Point(15, 30)
$chkRevoke.Width = 280
$grpGraph.Controls.Add($chkRevoke)

$chkGraphGroups = New-Object System.Windows.Forms.CheckBox
$chkGraphGroups.Text = (T "ChkGraphGroups")
$chkGraphGroups.Location = New-Object System.Drawing.Point(15, 60)
$chkGraphGroups.Width = 280
$grpGraph.Controls.Add($chkGraphGroups)

$chkLic = New-Object System.Windows.Forms.CheckBox
$chkLic.Text = (T "ChkLic")
$chkLic.Location = New-Object System.Drawing.Point(15, 90)
$chkLic.Width = 280
$grpGraph.Controls.Add($chkLic)

# Butonlar
$btnADOnly = New-Object System.Windows.Forms.Button
$btnADOnly.Text = (T "BtnADOnly")
$btnADOnly.Location = New-Object System.Drawing.Point(20, 320)
$btnADOnly.Size = New-Object System.Drawing.Size(280, 40)
$btnADOnly.BackColor = [System.Drawing.Color]::LightGreen
Set-RoundedCorners $btnADOnly 15
$Form.Controls.Add($btnADOnly)

$btnEXOOnly = New-Object System.Windows.Forms.Button
$btnEXOOnly.Text = (T "BtnEXOOnly")
$btnEXOOnly.Location = New-Object System.Drawing.Point(320, 320)
$btnEXOOnly.Size = New-Object System.Drawing.Size(320, 40)
$btnEXOOnly.BackColor = [System.Drawing.Color]::LightSkyBlue
Set-RoundedCorners $btnEXOOnly 15
$Form.Controls.Add($btnEXOOnly)

$btnGraphOnly = New-Object System.Windows.Forms.Button
$btnGraphOnly.Text = (T "BtnGraphOnly")
$btnGraphOnly.Location = New-Object System.Drawing.Point(660, 320)
$btnGraphOnly.Size = New-Object System.Drawing.Size(300, 40)
$btnGraphOnly.BackColor = [System.Drawing.Color]::Plum
Set-RoundedCorners $btnGraphOnly 15
$Form.Controls.Add($btnGraphOnly)

$btnRunAll = New-Object System.Windows.Forms.Button
$btnRunAll.Text = (T "BtnRunAll")
$btnRunAll.Location = New-Object System.Drawing.Point(20, 595)
$btnRunAll.Size = New-Object System.Drawing.Size(940, 40)
$btnRunAll.BackColor = [System.Drawing.Color]::LightSteelBlue
Set-RoundedCorners $btnRunAll 15
$Form.Controls.Add($btnRunAll)

function Set-ActionButtonsEnabled([bool]$enabled) {
    $btnADOnly.Enabled = $enabled
    $btnEXOOnly.Enabled = $enabled
    $btnGraphOnly.Enabled = $enabled
    $btnRunAll.Enabled = $enabled
}

function Set-UserLockedState([bool]$locked) {
    $txtSearch.Enabled = -not $locked
    $lstUsers.Enabled = -not $locked
    $btnLock.Text = if ($locked) { (T "BtnUnlock") } else { (T "BtnLock") }
    $btnLock.BackColor = if ($locked) { [System.Drawing.Color]::LightGreen } else { [System.Drawing.Color]::Khaki }
    Set-ActionButtonsEnabled $locked
}

# Moduller
function Ensure-Modules {
    try {
        Write-Log ">> Modules check..." "Yellow"
        Import-Module ActiveDirectory -ErrorAction Stop | Out-Null

        if (-not (Get-Module -ListAvailable ExchangeOnlineManagement)) {
            Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber
        }
        if (-not (Get-Module -ListAvailable Microsoft.Graph)) {
            Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
        }
        Write-Log "+ Modules ready." "Green"
    } catch {
        Write-Log "- Module error: $($_.Exception.Message)" "Red"
    }
}

function Disconnect-CloudSafely {
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
    try { Get-PSSession | Remove-PSSession -ErrorAction SilentlyContinue } catch {}
}

function Ensure-UserSelected {
    if (-not $Global:SelectedUser) { Write-Log "- No user selected." "Red"; return $false }
    return $true
}

# Privileged guard
function Test-IsPrivilegedUser {
    param([string]$SamAccountName)
    try {
        $domainSid = (Get-ADDomain -ErrorAction Stop).DomainSID.Value
        $criticalSids = @(
            "$domainSid-512", # Domain Admins
            "$domainSid-519", # Enterprise Admins
            "$domainSid-518", # Schema Admins
            "S-1-5-32-544"    # Builtin Administrators
        )

        $groups = Get-ADPrincipalGroupMembership -Identity $SamAccountName -ErrorAction Stop
        foreach ($g in $groups) {
            try {
                $sid = (Get-ADGroup -Identity $g.DistinguishedName -Properties objectSid -ErrorAction Stop).objectSid.Value
                if ($criticalSids -contains $sid) { return $true }
            } catch {
                if ($g.Name -in @("Domain Admins","Enterprise Admins","Schema Admins","Administrators")) { return $true }
            }
        }
        return $false
    } catch {
        return $true
    }
}

function Guard-PrivilegedUserOrAbort {
    if (-not $Global:SelectedUserId) { return $false }
    if (Test-IsPrivilegedUser -SamAccountName $Global:SelectedUserId) {
        [System.Windows.Forms.MessageBox]::Show((T "MsgPrivBody"), (T "MsgPrivTitle"),
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return $false
    }
    return $true
}

# Retry helper
function Invoke-WithRetry {
    param(
        [scriptblock]$Action,
        [int]$Retries = 3,
        [int]$DelaySeconds = 10,
        [string]$ActionName = "Operation"
    )
    for ($i=1; $i -le $Retries; $i++) {
        try {
            & $Action
            return $true
        } catch {
            if ($i -lt $Retries) {
                Write-Log "! $ActionName failed ($i/$Retries): $($_.Exception.Message) - waiting $DelaySeconds sec..." "Orange"
                Start-Sleep -Seconds $DelaySeconds
            } else {
                Write-Log "! WARN: $ActionName failed after $Retries attempts. Continuing. Error: $($_.Exception.Message)" "Orange"
                return $false
            }
        }
    }
}

# Purview PST (Search only)
# Purview PST (Search only)
function Start-OffboardPurviewMailboxSearch {
    param(
        [string]$Mailbox,
        [string]$SamAccountName
    )

    $searchName = "Offboard_{0}_{1}" -f $SamAccountName, (Get-Date -Format "yyyyMMdd_HHmm")

    try {
        New-ComplianceSearch -Name $searchName -ExchangeLocation $Mailbox -ContentMatchQuery "" -ErrorAction Stop | Out-Null
        Write-Log "+ Purview Search created: $searchName" "Green"

        Start-ComplianceSearch -Identity $searchName -ErrorAction Stop | Out-Null
        Write-Log "+ Purview Search started: $searchName" "Green"

        Write-Log (T "LogPurviewTitle") "Yellow"
        Write-Log (T "LogPurview1") "Yellow"
        Write-Log (T "LogPurview2") "Yellow"
        Write-Log ((T "LogPurview3") -f $searchName) "Yellow"
        Write-Log (T "LogPurview4") "Yellow"
        Write-Log (T "LogPurviewNote") "DarkGray"

        return $searchName
    } catch {
        Write-Log "! WARN: Purview Search error: $($_.Exception.Message)" "Orange"
        return $null
    }
}

# AD Organization cleanup (Manager + direct reports)
# AD Organization cleanup (Manager + direct reports)
function Invoke-ADOrganizationCleanup {
    param([string]$SamAccountName)

    try {
        $u = Get-ADUser -Identity $SamAccountName -Properties DistinguishedName, Manager -ErrorAction Stop
        $dn = $u.DistinguishedName

        $reportees = Get-ADUser -LDAPFilter "(manager=$dn)" -ErrorAction Stop
        foreach ($r in $reportees) {
            try { Set-ADUser -Identity $r.SamAccountName -Clear Manager -ErrorAction Stop } catch {}
        }

        if ($u.Manager) {
            try { Set-ADUser -Identity $SamAccountName -Clear Manager -ErrorAction Stop } catch {}
        }
    } catch {
        Write-Log "! WARN: Org cleanup error: $($_.Exception.Message)" "Orange"
    }
}

# AD On-Prem
function Invoke-ADOffboard {
    if (-not (Ensure-UserSelected)) { return }
    if (-not (Guard-PrivilegedUserOrAbort)) { return }

    $sam = $Global:SelectedUserId
    Write-Log "=== AD started: $sam ===" "Cyan"

    try {
        if ($chkADDisable.Checked) { Disable-ADAccount -Identity $sam; Write-Log "+ AD Disabled" "Green" }

        if ($chkADGroups.Checked) {
            $m = (Get-ADUser -Identity $sam -Properties MemberOf).MemberOf
            foreach ($g in $m) { try { Remove-ADGroupMember -Identity $g -Members $sam -Confirm:$false -ErrorAction Stop } catch {} }
            Write-Log "+ Removed from AD groups" "Green"
        }

        if ($chkADPwd.Checked) {
            Add-Type -AssemblyName System.Web
            $pwd = [System.Web.Security.Membership]::GeneratePassword(16,3)
            Set-ADAccountPassword -Identity $sam -Reset -NewPassword (ConvertTo-SecureString $pwd -AsPlainText -Force)
            Write-Log "+ Password reset (random)" "Green"
        }

        if ($chkADDesc.Checked) {
            $desc = "Offboarded - $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
            Set-ADUser -Identity $sam -Description $desc
            Write-Log "+ Description updated" "Green"
        }

        if ($chkADGal.Checked) {
            try { Set-ADUser -Identity $sam -Replace @{msExchHideFromAddressLists=$true} -ErrorAction Stop } catch {}
            try { Set-ADUser -Identity $sam -Clear showInAddressBook -ErrorAction Stop } catch {}
            Write-Log "+ GAL updated (Hide+Clear)" "Green"
        }

        if ($chkADOrg.Checked) {
            Invoke-ADOrganizationCleanup -SamAccountName $sam
            Write-Log "+ Org cleanup done (Manager/Reports)" "Green"
        }

        Write-Log "=== AD done ===" "Cyan"
    } catch {
        Write-Log "- AD error: $($_.Exception.Message)" "Red"
    }
}

# Exchange Online + Purview
function Invoke-EXOOffboard {
    if (-not (Ensure-UserSelected)) { return }
    if (-not (Guard-PrivilegedUserOrAbort)) { return }

    $smtp = $Global:SelectedSMTP
    Write-Log "=== EXO/Purview started: $smtp ===" "Cyan"
    Disconnect-CloudSafely

    try {
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        Write-Log "+ EXO connected" "Green"

        if ($chkShared.Checked) {
            Set-Mailbox -Identity $smtp -Type Shared -ErrorAction Stop
            Write-Log "+ Converted to shared mailbox" "Green"
        }

        if ($chkCal.Checked) {
            Invoke-WithRetry -Retries 3 -DelaySeconds 10 -ActionName "Remove-CalendarEvents" -Action {
                Remove-CalendarEvents -Identity $smtp -CancelOrganizedMeetings -QueryWindowInDays 365 -Confirm:$false -ErrorAction Stop
            } | Out-Null
        }

        if ($chkPurview.Checked) {
            try {
                Connect-IPPSSession -EnableSearchOnlySession -ErrorAction Stop
                Write-Log "+ Purview connected (IPPSSession)" "Green"
                [void](Start-OffboardPurviewMailboxSearch -Mailbox $smtp -SamAccountName $Global:SelectedUserId)
            } catch {
                Write-Log "! WARN: Purview error: $($_.Exception.Message)" "Orange"
            }
        }

        Write-Log "=== EXO/Purview done ===" "Cyan"
    } catch {
        Write-Log "- EXO error: $($_.Exception.Message)" "Red"
    } finally {
        Disconnect-CloudSafely
        Write-Log ">> Disconnect done" "DarkGray"
    }
}

# Graph / Entra ID
function Invoke-GraphOffboard {
    if (-not (Ensure-UserSelected)) { return }
    if (-not (Guard-PrivilegedUserOrAbort)) { return }

    $upnOrMail = $Global:SelectedSMTP
    Write-Log "=== Graph started: $upnOrMail ===" "Cyan"
    Disconnect-CloudSafely

    try {
        $scopes = @("User.ReadWrite.All","Directory.ReadWrite.All","Group.ReadWrite.All") | Select-Object -Unique
        Connect-MgGraph -Scopes $scopes -ErrorAction Stop | Out-Null
        Write-Log "+ Graph connected" "Green"

        $u = Get-MgUser -UserId $upnOrMail -ErrorAction Stop

        if ($chkRevoke.Checked) {
            Revoke-MgUserSignInSession -UserId $u.Id -ErrorAction Stop | Out-Null
            Write-Log "+ Revoke OK" "Green"
        }

        if ($chkGraphGroups.Checked) {
            $m = Get-MgUserMemberOf -UserId $u.Id -All -ErrorAction SilentlyContinue
            $groups = $m | Where-Object { $_.'@odata.type' -eq "#microsoft.graph.group" }
            foreach ($g in $groups) { try { Remove-MgGroupMemberByRef -GroupId $g.Id -DirectoryObjectId $u.Id -ErrorAction Stop } catch {} }
            Write-Log "+ Group removal attempted" "Green"
        }

        if ($chkLic.Checked) {
            try {
                $lic = Get-MgUserLicenseDetail -UserId $u.Id -ErrorAction Stop
                $skuIds = @($lic.SkuId) | Where-Object { $_ }
                if ($skuIds.Count -gt 0) {
                    Set-MgUserLicense -UserId $u.Id -AddLicenses @() -RemoveLicenses $skuIds -ErrorAction Stop | Out-Null
                    Write-Log "+ Licenses removed" "Green"
                } else {
                    Write-Log ">> No licenses found" "DarkGray"
                }
            } catch {
                Write-Log "! WARN: License removal error: $($_.Exception.Message)" "Orange"
            }
        }

        Write-Log "=== Graph done ===" "Cyan"
    } catch {
        Write-Log "- Graph error: $($_.Exception.Message)" "Red"
    } finally {
        Disconnect-CloudSafely
        Write-Log ">> Disconnect done" "DarkGray"
    }
}

# Kullanici arama
$searchTimer = New-Object System.Windows.Forms.Timer
$searchTimer.Interval = 350
$searchTimer.Add_Tick({
    $searchTimer.Stop()
    $q = $txtSearch.Text.Trim()
    $lstUsers.Items.Clear()
    $Global:SelectedUser = $null
    $Global:SelectedUserId = $null
    $Global:SelectedSMTP = $null
    $txtSMTP.Text = ""
    if ($q.Length -lt 2) { return }

    try {
        $filter = "Name -like '*$q*' -or SamAccountName -like '*$q*' -or UserPrincipalName -like '*$q*'"
        $res = Get-ADUser -Filter $filter -Properties EmailAddress, UserPrincipalName -ResultSetSize 10
        foreach ($u in $res) {
            $display = "{0} | {1} | {2}" -f $u.Name, $u.SamAccountName, $u.UserPrincipalName
            [void]$lstUsers.Items.Add($display)
        }
        if ($lstUsers.Items.Count -eq 0) { [void]$lstUsers.Items.Add("(No results)") }
    } catch {
        [void]$lstUsers.Items.Add("(Search error)")
        Write-Log "- AD search error: $($_.Exception.Message)" "Red"
    }
})
$txtSearch.Add_TextChanged({ $searchTimer.Stop(); $searchTimer.Start() })

$lstUsers.Add_SelectedIndexChanged({
    $sel = $lstUsers.SelectedItem
    if (-not $sel -or $sel -match "^\(") { return }
    $sam = ($sel -split "\|")[1].Trim()
    try {
        $Global:SelectedUser = Get-ADUser -Identity $sam -Properties EmailAddress, UserPrincipalName, DistinguishedName, Manager
        $Global:SelectedUserId = $Global:SelectedUser.SamAccountName
        $Global:SelectedSMTP = if ($Global:SelectedUser.EmailAddress) { $Global:SelectedUser.EmailAddress } else { $Global:SelectedUser.UserPrincipalName }
        $txtSMTP.Text = $Global:SelectedSMTP
        Write-Log ("+ {0}: {1} ({2})" -f (T "LogSelected"), $Global:SelectedUser.Name, $Global:SelectedUserId) "Green"
    } catch {
        Write-Log "- User load error: $($_.Exception.Message)" "Red"
    }
})

# Eventler
$btnModules.Add_Click({ Ensure-Modules })

$btnADConnect.Add_Click({
    try {
        Import-Module ActiveDirectory -ErrorAction Stop | Out-Null
        $btnADConnect.BackColor = [System.Drawing.Color]::LightGreen
        $btnADConnect.Text = (T "ADReady")
        Write-Log "+ AD module loaded." "Green"
    } catch {
        Write-Log "- AD module error: $($_.Exception.Message)" "Red"
    }
})

$btnLock.Add_Click({
    if (-not $Global:SelectedUser) { Write-Log "- Select a user first." "Red"; return }

    if ($txtSearch.Enabled) {
        if (-not (Guard-PrivilegedUserOrAbort)) { return }
        Set-UserLockedState $true
        Write-Log (">> {0}" -f (T "LogLocked")) "Cyan"
    } else {
        Set-UserLockedState $false
        Write-Log (">> {0}" -f (T "LogUnlocked")) "Cyan"
    }
})

$btnADOnly.Add_Click({ Invoke-ADOffboard })
$btnEXOOnly.Add_Click({ Invoke-EXOOffboard })
$btnGraphOnly.Add_Click({ Invoke-GraphOffboard })

$btnRunAll.Add_Click({
    if (-not (Ensure-UserSelected)) { return }
    if (-not (Guard-PrivilegedUserOrAbort)) { return }

    Write-Log (T "LogRunAllStart") "Yellow"
    Invoke-ADOffboard
    Invoke-EXOOffboard
    Invoke-GraphOffboard
    Write-Log (T "LogRunAllEnd") "Yellow"
})

# Dil degisince UI guncelle
function Apply-Language {
    $Form.Text = (T "AppTitle")
    $lblTitle.Text = $Form.Text

    $btnModules.Text = (T "BtnModules")
    $btnADConnect.Text = (T "BtnADConnect")
    $grpUser.Text = (T "GroupUser")
    $lblSearch.Text = (T "LblSearch")

    $grpAD.Text = (T "GroupAD")
    $grpEXO.Text = (T "GroupEXO")
    $grpGraph.Text = (T "GroupGraph")

    $btnADOnly.Text = (T "BtnADOnly")
    $btnEXOOnly.Text = (T "BtnEXOOnly")
    $btnGraphOnly.Text = (T "BtnGraphOnly")
    $btnRunAll.Text = (T "BtnRunAll")

    $chkADDisable.Text = (T "ChkADDisable")
    $chkADGroups.Text = (T "ChkADGroups")
    $chkADPwd.Text = (T "ChkADPwd")
    $chkADDesc.Text = (T "ChkADDesc")
    $chkADGal.Text = (T "ChkADGal")
    $chkADOrg.Text = (T "ChkADOrg")

    $chkCal.Text = (T "ChkCal")
    $chkShared.Text = (T "ChkShared")
    $chkPurview.Text = (T "ChkPurview")

    $chkRevoke.Text = (T "ChkRevoke")
    $chkGraphGroups.Text = (T "ChkGraphGroups")
    $chkLic.Text = (T "ChkLic")

    $btnLock.Text = if ($txtSearch.Enabled) { (T "BtnLock") } else { (T "BtnUnlock") }
}

$cmbLang.Add_SelectedIndexChanged({
    $Global:Lang = $cmbLang.SelectedItem.ToString()
    Apply-Language
})

# Baslangic
Set-ActionButtonsEnabled $false
Write-Log (T "LogReady") "Cyan"

$Form.Add_Shown({ Set-FormRoundedCorners -Form $Form -Radius 20; Apply-Language })
$Form.Add_Resize({ Set-FormRoundedCorners -Form $Form -Radius 20 })

[void]$Form.ShowDialog()