#Requires -RunAsAdministrator
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Threading

# --- Log file path ---
$LogFile = "C:\SCCM_Remediation_Log.txt"

# --- Ensure old log is archived ---
if (Test-Path $LogFile) {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    Rename-Item $LogFile "C:\SCCM_Remediation_Log_$timestamp.txt"
}

function Write-Log {
    param([string]$Message)
    $time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "[$time] $Message"
    Add-Content -Path $LogFile -Value $entry
}

function Show-Message {
    param(
        [string]$Message,
        [System.Windows.Forms.MessageBoxButtons]$Buttons = [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]$Icon = [System.Windows.Forms.MessageBoxIcon]::Information
    )
    [System.Windows.Forms.MessageBox]::Show($Message, "SCCM Remediation Tool", $Buttons, $Icon) | Out-Null
}

# --- Form ---
$form               = New-Object System.Windows.Forms.Form
$form.Text          = "SCCM Client Remediation Tool"
$form.Size          = New-Object System.Drawing.Size(700,720)
$form.StartPosition = "CenterScreen"
$form.TopMost       = $true

# --- Output Box ---
$outputBox                 = New-Object System.Windows.Forms.TextBox
$outputBox.Multiline       = $true
$outputBox.ScrollBars      = "Vertical"
$outputBox.Size            = New-Object System.Drawing.Size(650,280)
$outputBox.Location        = New-Object System.Drawing.Point(10,10)
$outputBox.ReadOnly        = $true
$outputBox.BackColor       = "Silver"
$outputBox.ForeColor       = "Black"
$outputBox.Font            = New-Object System.Drawing.Font("Consolas",10)
$form.Controls.Add($outputBox)

# --- GroupBox for Options ---
$groupBox                 = New-Object System.Windows.Forms.GroupBox
$groupBox.Text            = "Select Remediation Steps"
$groupBox.Size            = New-Object System.Drawing.Size(650,230)
$groupBox.Location        = New-Object System.Drawing.Point(10,300)
$form.Controls.Add($groupBox)

# --- "Select All" Checkbox ---
$selectAll = New-Object System.Windows.Forms.CheckBox
$selectAll.Text = "Select / Deselect All"
$selectAll.Location = New-Object System.Drawing.Point(15,30)
$selectAll.Size = New-Object System.Drawing.Size(200,20)
$groupBox.Controls.Add($selectAll)

# --- Step Checkboxes (keep a fixed order array) ---
$checkboxes = @{}
$labels = @(
    "Restart SCCM services",
    "Rebuild WMI repository",
    "Clear SCCM cache",
    "Reset Windows Update components",
    "Trigger SCCM client actions",
    "Reset WSUS registration"
)

for ($i=0; $i -lt $labels.Count; $i++) {
    $cb = New-Object System.Windows.Forms.CheckBox
    $cb.Text = $labels[$i]
    $cb.Location = New-Object System.Drawing.Point(35, (60 + ($i*25)))  # ✅ fixed
    $cb.Size = New-Object System.Drawing.Size(300,20)
    $groupBox.Controls.Add($cb)
    $checkboxes[$labels[$i]] = $cb
}

##for ($i=0; $i -lt $labels.Count; $i++) {
##    $cb = New-Object System.Windows.Forms.CheckBox
  #  $cb.Text = $labels[$i]
  #  $cb.Location = New-Object System.Drawing.Point(35,60 + ($i*25))  # indent
  #  $cb.Size = New-Object System.Drawing.Size(300,20)
  #  $groupBox.Controls.Add($cb)
  #  $checkboxes[$labels[$i]] = $cb
#}

# --- Select All Behavior ---
$selectAll.Add_CheckedChanged({
    foreach ($cb in $checkboxes.Values) {
        $cb.Checked = $selectAll.Checked
    }
})

# --- Status Label ---
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Status: Waiting to start..."
$statusLabel.AutoSize = $true
$statusLabel.Location = New-Object System.Drawing.Point(10,540)
$statusLabel.Font = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
$form.Controls.Add($statusLabel)

# --- Progress Bar ---
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10,565)
$progressBar.Size = New-Object System.Drawing.Size(650,25)
$progressBar.Style = 'Continuous'
$form.Controls.Add($progressBar)

# --- Buttons ---
$runButton                 = New-Object System.Windows.Forms.Button
$runButton.Text            = "Run Selected Remediation"
$runButton.Size            = New-Object System.Drawing.Size(250,40)
$runButton.Location        = New-Object System.Drawing.Point(10,600)
$form.Controls.Add($runButton)

$viewLogButton             = New-Object System.Windows.Forms.Button
$viewLogButton.Text        = "View Log"
$viewLogButton.Size        = New-Object System.Drawing.Size(100,40)
$viewLogButton.Location    = New-Object System.Drawing.Point(420,600)
$form.Controls.Add($viewLogButton)

$exitButton                 = New-Object System.Windows.Forms.Button
$exitButton.Text            = "Exit"
$exitButton.Size            = New-Object System.Drawing.Size(100,40)
$exitButton.Location        = New-Object System.Drawing.Point(560,600)
$form.Controls.Add($exitButton)

# --- Logging Helper (thread-safe with Invoke) ---
function Log-Output {
    param([string]$msg)
    Write-Log $msg  # log to file too
    if ($outputBox.InvokeRequired) {
        $outputBox.Invoke([Action]{ 
            $outputBox.AppendText("$msg`r`n")
            $outputBox.SelectionStart = $outputBox.Text.Length
            $outputBox.ScrollToCaret()
        })
    } else {
        $outputBox.AppendText("$msg`r`n")
        $outputBox.SelectionStart = $outputBox.Text.Length
        $outputBox.ScrollToCaret()
    }
}

# --- Status Update Helper ---
function Update-Status {
    param([string]$msg)
    Write-Log "STATUS: $msg"
    if ($statusLabel.InvokeRequired) {
        $statusLabel.Invoke([Action]{ $statusLabel.Text = "Status: $msg" })
    } else {
        $statusLabel.Text = "Status: $msg"
    }
}

# --- Progress Update Helper ---
function Update-ProgressBar {
    param([int]$value)
    if ($progressBar.InvokeRequired) {
        $progressBar.Invoke([Action]{ $progressBar.Value = $value })
    } else {
        $progressBar.Value = $value
    }
}

# --- Main Function to run in background runspace ---
$scriptBlock = {
    param($steps, $stepsTotal)

    # Safe cross-thread invocations (no quoting issues; use AddArgument)
    function Gui-Log { param([string]$msg) [PowerShell]::Create().AddScript('$global:LogAction.Invoke($args[0])').AddArgument($msg).Invoke() | Out-Null }
    function Gui-Progress { param([int]$val) [PowerShell]::Create().AddScript('$global:ProgressAction.Invoke($args[0])').AddArgument($val).Invoke() | Out-Null }
    function Gui-Status { param([string]$msg) [PowerShell]::Create().AddScript('$global:StatusAction.Invoke($args[0])').AddArgument($msg).Invoke() | Out-Null }

    $completed = 0
    try {
        Gui-Log "Starting remediation process..."

        foreach ($step in $steps) {
            $completed++
            Gui-Status ("Running step {0} of {1}: {2}" -f $completed, $stepsTotal, $step)

            switch ($step) {
                "Restart SCCM services" {
                    Gui-Log "Restarting SCCM related services..."
                    Restart-Service -Name ccmexec -Force -ErrorAction Stop
                    Restart-Service -Name wuauserv -Force -ErrorAction Stop
                    Restart-Service -Name bits -Force -ErrorAction Stop
                }
                "Rebuild WMI repository" {
                    Gui-Log "Rebuilding WMI repository..."
                    Stop-Service -Name winmgmt -Force -ErrorAction Stop
                    if (Test-Path "$env:SystemRoot\System32\wbem\Repository\") {
                        Remove-Item "$env:SystemRoot\System32\wbem\Repository\" -Recurse -Force -ErrorAction Stop
                    }
                    Start-Service -Name winmgmt -ErrorAction Stop
                    winmgmt /resetrepository | Out-Null
                }
                "Clear SCCM cache" {
                    Gui-Log "Clearing SCCM cache..."
                    Get-Service -Name ccmexec | Stop-Service -Force -ErrorAction Stop
                    if (Test-Path "$env:Windir\CCM\Cache\") {
                        Remove-Item "$env:Windir\CCM\Cache\*" -Recurse -Force -ErrorAction Stop
                    }
                    Start-Service -Name ccmexec -ErrorAction Stop
                }

                                "Reset Windows Update components" {
                    Gui-Log "Resetting Windows Update components..."
                    Get-Service -Name wuauserv,bits | Stop-Service -Force -ErrorAction Stop
                    Rename-Item "$env:SystemRoot\SoftwareDistribution" "SoftwareDistribution.old" -Force -ErrorAction Stop
                    Rename-Item "$env:SystemRoot\System32\catroot2" "catroot2.old" -Force -ErrorAction Stop
                    Start-Service -Name bits -ErrorAction Stop
                    Start-Service -Name wuauserv -ErrorAction Stop
                }

                    "Reset Windows Update components" {
                    Gui-Log "Resetting Windows Update components..."
                    Get-Service -Name wuauserv,bits | Stop-Service -Force -ErrorAction Stop
                    Rename-Item "$env:SystemRoot\SoftwareDistribution" "SoftwareDistribution.old" -Force -ErrorAction Stop
                    Rename-Item "$env:SystemRoot\System32\catroot2" "catroot2.old" -Force -ErrorAction Stop
                    Start-Service -Name bits -ErrorAction Stop
                    Start-Service -Name wuauserv -ErrorAction Stop
                }

                "Trigger SCCM client actions" {
                    Gui-Log "Triggering SCCM client actions..."
                    & "$env:Windir\CCM\ccmexec.exe" /forcepolicy
                    & "$env:Windir\CCM\ccmexec.exe" /resetpolicy
                    & "$env:Windir\CCM\ccmexec.exe" /mp:your-sccm-server.fqdn
                    & "$env:Windir\CCM\ccmexec.exe" /MachinePolicy
                    & "$env:Windir\CCM\ccmexec.exe" /RequestMachinePolicy
                    & "$env:Windir\CCM\ccmexec.exe" /eval
                    & "$env:Windir\CCM\ccmexec.exe" /SoftwareUpdateScan
                }

            $percent = [math]::Round(($completed / $stepsTotal) * 100)
            Gui-Progress $percent
        }

        Gui-Status "All tasks completed ✅"
        Gui-Log "Remediation completed ✅"
        [PowerShell]::Create().AddScript("[System.Windows.Forms.MessageBox]::Show('Remediation completed. Please check updates again. A reboot may be required.`nLog saved to: $using:LogFile','SCCM Remediation Tool')").Invoke() | Out-Null
    }
    catch {
        Gui-Status "Error occurred ❌"
        Gui-Log "❌ Error: $_"
        [PowerShell]::Create().AddScript("[System.Windows.Forms.MessageBox]::Show('Error: $_`nSee log file: $using:LogFile','SCCM Remediation Tool',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)").Invoke() | Out-Null
    }
}

# --- Button Events ---
$runButton.Add_Click({
    # Build an ORDERED list of selected steps so numbering is stable
    $selectedStepsOrdered = @($labels | Where-Object { $checkboxes[$_].Checked })

    $stepsTotal = $selectedStepsOrdered.Count
    if ($stepsTotal -eq 0) {
        Show-Message "Please select at least one remediation step." -Buttons OK -Icon Warning
        return
    }

    $progressBar.Value = 0
    Update-Status "Starting..."
    $runButton.Enabled = $false

    $global:LogAction     = { param($msg) Log-Output $msg }
    $global:ProgressAction= { param($val) Update-ProgressBar $val }
    $global:StatusAction  = { param($msg) Update-Status $msg }

    $ps = [PowerShell]::Create().AddScript($scriptBlock).AddArgument($selectedStepsOrdered).AddArgument($stepsTotal)

    # Re-enable button when the runspace finishes
    $async = $ps.BeginInvoke()
    Register-ObjectEvent -InputObject $async -EventName 'Completed' -Action {
        # marshal back to GUI thread
        $form.BeginInvoke([Action]{ $runButton.Enabled = $true }) | Out-Null
        Unregister-Event -SourceIdentifier $event.SourceIdentifier
        $event.Sender.Dispose()
    } | Out-Null
})

$viewLogButton.Add_Click({
    if (Test-Path $LogFile) {
        Start-Process notepad.exe $LogFile
    } else {
        Show-Message "No log file found yet." -Buttons OK -Icon Warning
    }
})

$exitButton.Add_Click({ $form.Close() })

# --- Show Form ---
[void]$form.ShowDialog()
