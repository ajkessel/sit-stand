#Requires -Version 5.1
<#
.SYNOPSIS
    This script will remind you to stand up and sit down on a defined schedule, with a countdown in the systray.
.DESCRIPTION
    You shouldn't sit all day! This script will remind you to sit and stand on a defined schedule.
.PARAMETER sit
    Time to sit before being reminded to sit in minutes (default 60)
.PARAMETER stand
    Time to stand before being reminded to sit in minutes (default 60)
.PARAMETER font
    Font in systray
.PARAMETER fg
    Foreground color in systray
.PARAMETER bg
    Background color in systray
.PARAMETER auto
    Start automatically on login
.PARAMETER help
    Display this help message
.EXAMPLE
  > sit-stand -sit 45 -stand 15
#>
Param(
  [Parameter(Mandatory = $false, HelpMessage = "Time to sit in minutes (default 60)")][int]$sit = 0,
  [Parameter(Mandatory = $false, HelpMessage = "Time to stand in minutes (default 60)")][int]$stand = 0,
  [Parameter(Mandatory = $false, HelpMessage = "Path to sitting reminder sound file")][string]$sitsoundfile = "",
  [Parameter(Mandatory = $false, HelpMessage = "Path to standing reminder sound file")][string]$standsoundfile = "",
  [Parameter(Mandatory = $false, HelpMessage = "Systray Font (default Calibri)")][string]$font = "",
  [Parameter(Mandatory = $false, HelpMessage = "Systray Foreground Color (default White)")][string]$fg = "",
  [Parameter(Mandatory = $false, HelpMessage = "Systray Background Color (default DarkBlue)")][string]$bg = "",
  [Parameter(Mandatory = $false, HelpMessage = "Respect focus assist")][switch]$focus,
  [Parameter(Mandatory = $false, HelpMessage = "Start automatically on login")][switch]$auto,
  [Parameter(Mandatory = $false, HelpMessage = "Display this help message")][switch]$help,
  [Parameter(Mandatory = $false, HelpMessage = "Uninstall")][switch]$uninstall
)
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName Microsoft.VisualBasic
$global:RegPath = "HKCU:\Software\Sit-Stand"
if ($help) {
  write-output("Sit-Stand Reminder`n`nYou shouldn't sit all day! This script will remind you to sit and stand on a defined schedule.`nParameters:`n-sit [###] Time to sit before being remanded to stand in minutes (default 60)`n-stand [###] Time to stand before being remanded to sit in minutes (default 60)`n-font typeface in systray`n-fg foreground color in systray`n-bg background color in systray`n-auto start automatically on login`n-uninstall remove all traces of this application`n-help Display this help message`n`nCopyright (2024) Adam J. Kessel - MIT License")
  exit
}
if ($uninstall) {
  if ( test-path $global:RegPath ) { Remove-Item $global:RegPath }
  write-output("Uninstall complete. You can delete this application now.")
  exit
}
if ( test-path $global:RegPath ) {
  if ( -not $sit -and ($x = (get-item -path $global:RegPath).GetValue("sit"))) {
    $sit = $x
  }
  if ( -not $stand -and ($x = (get-item -path $global:RegPath).GetValue("stand"))) {
    $stand = $x
  }
  if ( -not $font -and ($x = (get-item -path $global:RegPath).GetValue("font"))) {
    $font = $x
  }
  if ( -not $focus -and ($x = (get-item -path $global:RegPath).GetValue("focus"))) {
    $focus = if ( $x ) { $true } else { $false }
  }
  if ( -not $fg -and ($x = (get-item -path $global:RegPath).GetValue("fg"))) {
    $fg = $x
  }
  if ( -not $bg -and ($x = (get-item -path $global:RegPath).GetValue("bg"))) {
    $bg = $x
  }
  if ( -not $sitsoundfile -and ($x = (get-item -path $global:RegPath).GetValue("sitsound"))) {
    $sitsoundfile = $x
  }
  if ( -not $standsoundfile -and ($x = (get-item -path $global:RegPath).GetValue("standsound"))) {
    $standsoundfile = $x
  }
}
if ( -not $font ) { $font = "Calibri" }
if ( -not $fg ) { $fg = "White" }
if ( -not $bg ) { $bg = "DarkBlue" }
if ( -not $sitsoundfile ) { $sitsoundfile = "c:\windows\media\Windows Balloon.wav" }
if ( -not $standsoundfile ) { $standsoundfile = "c:\windows\media\Windows Default.wav" }
if ( -not ( test-path $sitsoundfile ) ) { $sitsoundfile = "silent" }
if ( -not ( test-path $standsoundfile ) ) { $standsoundfile = "silent" }
if ($sit -lt 1) { $sit = 60 }
if ($stand -lt 1) { $stand = 60 }
if ($sit -gt 999) { $sit = 999 }
if ($stand -gt 999) { $stand = 999 }
if ( -not ( test-path $global:RegPath )) {
  $null = New-Item $global:RegPath
}
$global:SitTime = $sit
$global:StandTime = $stand
$global:TrayFont = $font
$global:TrayColorBg = $bg
$global:TrayColorfg = $fg
$global:objNotifyIcon = New-Object System.Windows.Forms.NotifyIcon
$global:contextMenu = [System.Windows.Forms.ContextMenuStrip]::new()
$global:SitSound=New-Object System.Media.SoundPlayer
$global:SitSound.SoundLocation=$SitSoundFile
$global:StandSound=New-Object System.Media.SoundPlayer
$global:StandSound.SoundLocation=$StandSoundFile
$global:focus = $focus
If ($auto) {
  $global:autostart = $true
}
else {
  $global:autostart = ((get-item -path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run').GetValue("Sit-Stand"))
}

$MainFunction = {
  $global:remaining = $global:sittime
  $global:sitting = $true
  $global:done = $false
  UpdateRegistry
  SetTray
  while ( -not $global:done -and $y -ne 'Cancel' ) {
    UpdateIcon
    while ( -not $global:done -and $global:remaining -gt 0 ) {
      $global:counter = 600
      while ( $global:counter -gt 0 -and -not $global:done ) {
        [System.Windows.Forms.Application]::DoEvents()
        start-sleep -milliseconds 100
        $global:counter -= 1
      }
      if ( $global:remaining -gt 0 ) {
        $global:remaining -= 1
        UpdateIcon
      }
    }
    if ( -not $global:done -and $global:remaining -gt -1 ) {
      if ( $global:sitting ) {
        if ( DisturbOk ) {
        if ( $StandSound.Location -ne "silent" ) { $StandSound.PlaySync() }
        $y = [Microsoft.VisualBasic.Interaction]::MsgBox('Stand Up!','OKCancel,SystemModal,Information', 'Stand Reminder')
        }
      }
      else {
        if ( DisturbOk ) {
        if ( $SitSound.Location -ne "silent" ) { $SitSound.PlaySync() }
        $y = [Microsoft.VisualBasic.Interaction]::MsgBox('Sit down.','OKCancel,SystemModal,Information', 'Stand Reminder')
        }
      }
      $global:sitting = -not $global:sitting
    }
    if ($global:sitting) {
      $global:remaining = $global:sittime
    }
    else {
      $global:remaining = $global:standtime
    }
  }
  $global:objNotifyIcon.Dispose()
}

function SetTray {
  Set-Variable -Name ($contextMenu, $objNotifyIcon) -Scope Global
  $objNotifyIcon.Visible = $True
  $objNotifyIcon.Text = "Stand Reminder"
  $objNotifyIcon.Visible = $True
  $menuItemStand = New-Object System.Windows.Forms.ToolStripMenuItem
  $menuItemStand.Text = "Stand"
  $menuItemSit = New-Object System.Windows.Forms.ToolStripMenuItem
  $menuItemSit.Text = "Sit"
  $menuItemSit.Checked = $true
  $menuItemChange = New-Object System.Windows.Forms.ToolStripMenuItem
  $menuItemChange.Text = "Change"
  $menuItemChangeSit = New-Object System.Windows.Forms.ToolStripMenuItem
  $menuItemChangeSit.Text = ("Sit Time (" + $Global:SitTime + ")")
  $menuItemChangeStand = New-Object System.Windows.Forms.ToolStripMenuItem
  $menuItemChangeStand.Text = ("Stand Time (" + $Global:StandTime + ")")
  $menuItemChangeFont = New-Object System.Windows.Forms.ToolStripMenuItem
  $menuItemChangeFont.Text = "Font"
  $menuItemChangeFg = New-Object System.Windows.Forms.ToolStripMenuItem
  $menuItemChangeFg.Text = "Foreground"
  $menuItemChangeBg = New-Object System.Windows.Forms.ToolStripMenuItem
  $menuItemChangeBg.Text = "Background"
  $menuItemChangeAuto = New-Object System.Windows.Forms.ToolStripMenuItem
  $menuItemChangeAuto.Text = "Start automatically"
  $menuItemChangeAuto.Add_Click({
      $global:autostart = (-not $global:autostart)
      UpdateIcon
      UpdateRegistry
    })
  $menuItemChangeFocus = New-Object System.Windows.Forms.ToolStripMenuItem
  $menuItemChangeFocus.Text = "Respect focus assist"
  $menuItemChangeFocus.Add_Click({
      $global:focus = (-not $global:focus)
      UpdateIcon
      UpdateRegistry
    })
  $menuItemChangeSitSound = New-Object System.Windows.Forms.ToolStripMenuItem
  $menuItemChangeSitSound.Text = "Sit reminder sound file"
  $menuItemChangeSitSound.Add_Click({
      $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = [Environment]::GetFolderPath('Desktop') }
      if ( test-path c:\windows\media ) {
      $FileBrowser.InitialDirectory = "c:\windows\media"
      } else {
      $FileBrowser.InitialDirectory = "[Environment]::GetFolderPath('Desktop')"
      }
      $a = $FileBrowser.ShowDialog()
      if ( $a -eq 'Cancel' ) { $global:SitSound.SoundLocation = "silent" } else {
      if ( test-path $FileBrowser.Filename ) {
      $global:SitSound.SoundLocation=($FileBrowser.Filename)
      } else {
      $global:SitSound.SoundLocation="silent"
      }
      }
      UpdateIcon
      UpdateRegistry
      })
  $menuItemChangeStandSound = New-Object System.Windows.Forms.ToolStripMenuItem
  $menuItemChangeStandSound.Text = "Stand reminder sound file"
  $menuItemChangeStandSound.Add_Click({
      $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = [Environment]::GetFolderPath('Desktop') }
      if ( test-path c:\windows\media ) {
      $FileBrowser.InitialDirectory = "c:\windows\media"
      } else {
      $FileBrowser.InitialDirectory = "[Environment]::GetFolderPath('Desktop')"
      }
      $a = $FileBrowser.ShowDialog()
      if ( $a -eq 'Cancel' ) { $global:StandSound.SoundLocation = "silent" } else {
      if ( test-path $FileBrowser.Filename ) {
      $global:StandSound.SoundLocation=($FileBrowser.Filename)
      } else {
      $global:StandSound.SoundLocation="silent"
      }
      }
      UpdateIcon
      UpdateRegistry
      })
  
  @("Arial", "Calibri", "Comic Sans MS", "Courier New", "Georgia", "Impact", "Lucida Console", "Palatino", "Tahoma", "Times New Roman", "Trebuchet MS", "Verdana", "MS Sans Serif", "MS Serif") | ForEach {
    $x = New-Object System.Windows.Forms.ToolStripMenuItem
    $x.Text = $_
    $x.Add_Click( {
        param(
          [System.Object] $selected
        )
        $Global:TrayFont = ($selected.PSObject.Copy()).text
        UpdateRegistry
        UpdateIcon
      } )
    $null = $menuItemChangeFont.DropDownItems.Add($x)
  }
  @([System.Enum]::GetValues('ConsoleColor')) | ForEach {
    $fg = New-Object System.Windows.Forms.ToolStripMenuItem
    $bg = New-Object System.Windows.Forms.ToolStripMenuItem
    $fg.Text = $_  
    $bg.Text = $_
    $fg.Add_Click( {
        param( [System.Object] $selected)
        $global:TrayColorFg = ($selected.PSObject.Copy()).text
        UpdateRegistry
        UpdateIcon
      } )
    $bg.Add_Click( {
        param( [System.Object] $selected)
        $global:TrayColorBg = ($selected.PSObject.Copy()).text
        UpdateRegistry
        UpdateIcon
      } )
    $null = $menuItemChangeFg.DropDownItems.Add($fg)
    $null = $menuItemChangeBg.DropDownItems.Add($bg) 
  }
  $menuItemChangeSit.add_Click({
      $newTime = [int](ChangeTime "sit")
      if ($newTime -gt 0 -and $newTime -lt 1000) {
        $Global:sitTime = $newTime
        $global:counter = -1
        $global:remaining = -1
        $contextMenu.Items[2].DropDownItems[0].Text = ("Sit Time (" + $Global:SitTime + ")")
        UpdateRegistry
      }
    })
  $menuItemChangeStand.add_Click({
      $newTime = [int](ChangeTime "stand")
      if ($newTime -gt 0 -and $newTime -lt 1000) {
        $Global:standTime = $newTime
        $global:counter = -1
        $global:remaining = -1
        $contextMenu.Items[2].DropDownItems[1].Text = ("Stand Time (" + $Global:StandTime + ")")
        UpdateRegistry
      }
    })
  @($menuItemChangeSit, $menuItemChangeStand, $menuItemChangeFont, $menuItemChangeFg, $menuItemChangeBg, $menuItemChangeAuto, $menuItemChangeFocus, $menuItemChangeSitSound, $menuItemChangeStandSound) | ForEach {
    $null = $menuItemChange.DropDownItems.Add($_)
  }
  $menuItemExit = [System.Windows.Forms.ToolStripMenuItem]::new()
  $menuItemExit.Text = "Exit"
  $menuItemSit.add_Click({ 
      $global:counter = -1
      $global:remaining = -1
      $global:sitting = $true
      $contextMenu.Items[0].Checked = $true
      $contextMenu.Items[1].Checked = $false
    })
  $menuItemStand.add_Click({ 
      $global:counter = -1
      $global:remaining = -1
      $global:sitting = $false
      $contextMenu.Items[0].Checked = $false
      $contextMenu.Items[1].Checked = $true
    })
  $menuItemExit.add_Click({ 
      $global:done = $true 
    })
  $null = $contextMenu.Items.Add($menuItemSit)
  $null = $contextMenu.Items.Add($menuItemStand)
  $null = $contextMenu.Items.Add($menuItemChange)
  $null = $contextMenu.Items.Add($menuItemExit)
  $objNotifyIcon.ContextMenuStrip = $contextmenu
}

function UpdateIcon {
  $objNotifyIcon.Icon = GetIcon $global:remaining $global:sitting
  $contextMenu.Items[0].Text = ("Sit" + (& { If ($global:sitting) { " (" + $Global:Remaining + ")" } }))
  $contextMenu.Items[1].Text = ("Stand" + (& { If (-not $global:sitting) { " (" + $Global:Remaining + ")" } }))
  $contextMenu.Items[0].Checked = ( $global:sitting) 
  $contextMenu.Items[1].Checked = ( -not $global:sitting) 
  foreach ($x in $contextMenu.Items[2].DropDownItems[2].DropDownItems) {
    $x.checked = ($x.text -eq $global:TrayFont)
  }
  foreach ($x in $contextMenu.Items[2].DropDownItems[3].DropDownItems) {
    $x.checked = ($x.text -eq $global:TrayColorFg)
  }
  foreach ($x in $contextMenu.Items[2].DropDownItems[4].DropDownItems) {
    $x.checked = ($x.text -eq $global:TrayColorBg)
  }
  $contextMenu.Items[2].DropDownItems[6].Checked = $global:focus
  if ( $global:SitSound.SoundLocation -eq 'silent' ) {
  $contextMenu.Items[2].DropDownItems[7].Text = "Sit sound (silent)"
  } else {
  $contextMenu.Items[2].DropDownItems[7].Text = ("Sit sound (" + [System.IO.Path]::GetFileNameWithoutExtension($global:SitSound.SoundLocation) + ")")
  }
  if ( $global:StandSound.SoundLocation -eq 'silent' ) {
  $contextMenu.Items[2].DropDownItems[8].Text = "Stand sound (silent)"
  } else {
  $contextMenu.Items[2].DropDownItems[8].Text = ("Stand sound (" + [System.IO.Path]::GetFileNameWithoutExtension($global:StandSound.SoundLocation) + ")")
  }
}

function ChangeTime {
  param([string]$x)
  $title = 'New time'
  $msg = 'Enter new ' + $x + ' time interval:'
  $text = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
  return($text -replace '[^0-9]', '')
}

function UpdateRegistry {
  Set-ItemProperty -Path $global:RegPath -Name 'sit' -Value $global:SitTime -Type 'dword'
  Set-ItemProperty -Path $global:RegPath -Name 'stand' -Value $global:StandTime -Type 'dword'
  Set-ItemProperty -Path $global:RegPath -Name 'font' -Value $global:TrayFont -Type 'string'
  Set-ItemProperty -Path $global:RegPath -Name 'fg' -Value $Global:TrayColorFg -Type 'string'
  Set-ItemProperty -Path $global:RegPath -Name 'bg' -Value $global:TrayColorBg -Type 'string'
  Set-ItemProperty -Path $global:RegPath -Name 'sitsound' -Value $global:SitSound.SoundLocation -Type 'string'
  Set-ItemProperty -Path $global:RegPath -Name 'standsound' -Value $global:StandSound.SoundLocation -Type 'string'
  If ($global:focus) {
     Set-ItemProperty -Path $global:RegPath -Name 'focus' -Value "1" -Type 'dword'
  } else {
     Set-ItemProperty -Path $global:RegPath -Name 'focus' -Value "0" -Type 'dword'
  }
  If ($global:autostart) {
    $runTask = If ($PSCommandPath) { $PSCommandPath } else { (get-process -id $pid).path }
    Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run' -Name "Sit-Stand" -Value $runTask
  }
  elseif ((Get-Item -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run').GetValue("Sit-Stand")) {
    Remove-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run' -Name "Sit-Stand"
  }
}
# GetIcon 
Function GetIcon {
  param (
    [Parameter(Mandatory = $true)][int]$number,
    [Parameter(Mandatory = $false)][bool]$reverse = $false
  )
  <#
        .SYNOPSIS
        Creates a graphical icon based on a number for use in the systray.

        .PARAMETER Number
        Number to convert to an image.

        .PARAMETER Reverse
        Invert colors (Boolean, optional)

        .OUTPUTS
        System.Drawing.Icon. An icon object that can be added to the systray.
      #>
  $bmp = new-object System.Drawing.Bitmap 128, 128
  $font = new-object System.Drawing.Font $Global:TrayFont, ([math]::floor(100 / ([string]$number).length))
  if ($reverse) {
    $bg = new-object System.Drawing.SolidBrush $Global:TrayColorFg
    $fg = new-object System.Drawing.SolidBrush $Global:TrayColorBg
  }
  else {
    $fg = new-object System.Drawing.SolidBrush $Global:TrayColorFg
    $bg = new-object System.Drawing.SolidBrush $Global:TrayColorBg
  }
  $graphics = [System.Drawing.Graphics]::FromImage($bmp)
  $graphics.FillRectangle($fg, 0, 0, $bmp.Width, $bmp.Height)
  $Format = [System.Drawing.StringFormat]::GenericDefault
  $Format.Alignment = [System.Drawing.StringAlignment]::Center
  $Format.LineAlignment = [System.Drawing.StringAlignment]::Center
  $Rectangle = [System.Drawing.RectangleF]::FromLTRB(0, 0, 128, 128)
  $graphics.DrawString($number, $Font, $bg, $rectangle, $format)
  $graphics.Dispose()
  Return [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($bmp).GetHIcon()))
}

# experimental setting to check "focus assist"
function Get-QuietHoursProfile {
    try {
    $rawData = Get-ItemPropertyValue 'HKCU:\Software\Microsoft\Windows\CurrentVersion\CloudStore\Store\Cache\DefaultAccount\$$windows.data.notifications.quiethourssettings\Current' Data
    $string = [Text.Encoding]::Unicode.GetString($rawData[0x1a..($rawData.length-1)])
    
    if ($string -notmatch 'Microsoft\.QuietHoursProfile\.(\w+)') { return ""}
    return $matches[1]
    } catch { return "" }
}

function DisturbOk {
   $x = Get-QuietHoursProfile
   if ( $x -match 'Only' -and $global:focus ) {
     return $false
   } else { return $true }
}

& $MainFunction
