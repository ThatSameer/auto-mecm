Clear-Host
hostname

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic
[System.Windows.Forms.Application]::EnableVisualStyles()

# Import settings.conf
$confFile = "$($PSScriptRoot | Split-Path)\settings.conf"
if (-not(Test-Path -Path $confFile -PathType Leaf)) {
    Write-Host -ForegroundColor Red 'ERROR: settings.conf file does not exist! Exiting...'
    Pause
    break 
}

foreach ($i in $(Get-Content $confFile)) {
    Set-Variable -Name $i.split('=')[0] -Value $i.split('=', 2)[1]
}

$logDir = "$LOGLOCATION\$env:USERNAME"
[Void][System.IO.Directory]::CreateDirectory($logDir)

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]
        $Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet('INFO', 'WARN', 'SUCCESS', 'ERROR', 'FATAL', 'DEBUG')]
        [string]
        $Level = 'INFO',

        [Parameter(Mandatory = $false)]
        [string]
        $LogFile = ('FileSystem::' + "$logDir\Auto-MECM_Log.log")
    )

    $consoleDate = (Get-Date).toString('HH:mm:ss')
    $consoleLog = Switch ($Level) {
        'WARN' { Write-Host "$consoleDate $Level $Message" -ForegroundColor Yellow }
        'SUCCESS' { Write-Host "$consoleDate $Level $Message" -ForegroundColor Green }
        'ERROR' { Write-Host "$consoleDate $Level $Message" -ForegroundColor Red }
        'FATAL' { Write-Host "$consoleDate $Level $Message" -ForegroundColor Red }
        'DEBUG' { Add-Content $LogFile -Value 'DEBUG' }
        default { Write-Host "$consoleDate $Level $Message" }
    }
    $stamp = (Get-Date).toString('yyyy/MM/dd HH:mm:ss')
    $line = "$stamp $Level $Message"

    if ($LogFile) {
        Add-Content $LogFile -Value $line
        $consoleLog
    }
    else {
        $consoleLog
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Auto MECM'
$form.Size = New-Object System.Drawing.Size(300, 250)
$form.StartPosition = 'CenterScreen'

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(75, 170)
$okButton.Size = New-Object System.Drawing.Size(75, 23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(150, 170)
$cancelButton.Size = New-Object System.Drawing.Size(75, 23)
$cancelButton.Text = 'Cancel'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10, 20)
$label.Size = New-Object System.Drawing.Size(280, 20)
$label.Text = 'Please select a tool to run:'
$form.Controls.Add($label)

$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10, 40)
$listBox.Size = New-Object System.Drawing.Size(260, 20)
$listBox.Height = 130

[void] $listBox.Items.Add('1) Update Maintenance Window')
[void] $listBox.Items.Add('2) Find Server Collection')
[void] $listBox.Items.Add('3) Add Server to Collection')
[void] $listBox.Items.Add('4) Remove Server from Collection')

$form.Controls.Add($listBox)
$form.Topmost = $true
$form.MaximizeBox = $false
$form.MinimizeBox = $false

$form.Add_Shown({ $form.Activate() })
$result = $form.ShowDialog()

if (($result -eq [System.Windows.Forms.DialogResult]::OK) -and ($null -ne $listBox.SelectedItem)) {
    Write-Log -Message "You have selected $($listBox.SelectedItem)" -Level INFO
    $toolChoice = $listBox.SelectedItem
}
else { 
    Write-Log -Message 'Selection cancelled by user. Exiting...' -Level ERROR
    Pause
    break
}

#Import MECM module
Write-Log -Message 'Importing the MECM module. Please wait...' -Level INFO
try {
    Import-Module (Join-Path $(Split-Path $env:SMS_ADMIN_UI_PATH) ConfigurationManager.psd1)
    New-PSDrive -Name $SITECODE -PSProvider 'CMSite' -Root $SITESERVER -Description "($SITECODE - $SITESERVER, Primary Site Server)" | Out-Null
    Set-Location -Path "$($SITECODE):\"
}
catch {
    Write-Log -Message 'Failed to import the MECM module. Exiting.' -Level ERROR
    Pause
    break
}

#Invoke script based on selection
Switch -Regex ($toolChoice) {
    '^1\)' { Invoke-Expression "$PSScriptRoot\Options\Set-MaintenanceWindow.ps1" }
    '^2\)' { Invoke-Expression "$PSScriptRoot\Options\Find-ServerCollection.ps1" }
    '^3\)' { Invoke-Expression "$PSScriptRoot\Options\Add-ServerToCollection.ps1" }
    '^4\)' { Invoke-Expression "$PSScriptRoot\Options\Remove-ServerFromCollection.ps1" }
    Default { 
        Write-Log -Message 'Invalid option selected. Exiting.' -Level ERROR
        break
    }
}

Write-Host "`nThank you for using Auto-MECM on Citrix Storefront. Exiting..."
Pause
break