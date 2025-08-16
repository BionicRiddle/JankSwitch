param(
    [switch]$ListConfig,
    [switch]$SaveConfig
)

# Config file path
$ScriptDir = if ($PSScriptRoot) {
    $PSScriptRoot  # Running as .ps1
} else {
    Split-Path -Parent ([System.Reflection.Assembly]::GetEntryAssembly().Location) # Running as .exe
}

$ConfigPath = Join-Path $ScriptDir "audio.conf"

function Get-AudioDevice {
    $result = & (Get-Command -Name Get-AudioDevice -CommandType Cmdlet) @args
    return $result
}

function Set-AudioDevice {
    $null = & (Get-Command -Name Set-AudioDevice -CommandType Cmdlet) @args
}

function Get-CurrentAudioConfig {
    $playbackDef = Get-AudioDevice -Playback
    $recordingDef = Get-AudioDevice -Recording
    $playbackComm = Get-AudioDevice -PlaybackCommunication
    $recordingComm = Get-AudioDevice -RecordingCommunication

    # Fallback if no comm device set
    if (-not $playbackComm) { $playbackComm = $playbackDef }
    if (-not $recordingComm) { $recordingComm = $recordingDef }

    return @{
        DefaultPlaybackName = $playbackDef.Name
        DefaultCommPlaybackName = $playbackComm.Name
        DefaultRecordingName = $recordingDef.Name
        DefaultCommRecordingName = $recordingComm.Name
    }
}

# Ensure AudioDeviceCmdlets module is installed
if (-not (Get-Module -ListAvailable -Name AudioDeviceCmdlets)) {
    Write-Host "AudioDeviceCmdlets module not found. Please install it using the following command:`n`nInstall-Module -Name AudioDeviceCmdlets -Scope CurrentUser`n"
    exit 1
}
Import-Module AudioDeviceCmdlets

if ($SaveConfig) {
    $config = Get-CurrentAudioConfig
    
    try {
        $config | ConvertTo-Json | Set-Content -Path $ConfigPath

        Write-Host ("`nConfiguration saved to $ConfigPath`n`n" +
            "`$DefaultPlaybackName       = `"$($config.DefaultPlaybackName)`"  # Multimedia Playback`n" +
            "`$DefaultCommPlaybackName   = `"$($config.DefaultCommPlaybackName)`"  # Communications Playback`n" +
            "`$DefaultRecordingName      = `"$($config.DefaultRecordingName)`"  # Multimedia Recording`n" +
            "`$DefaultCommRecordingName  = `"$($config.DefaultCommRecordingName)`"  # Communications Recording")
        
    }
    catch {
        Write-Error "Failed to save configuration to $ConfigPath`n`nError: $($_.Exception.Message)`n"
    }
    
    exit
}

if ($ListConfig) {
    $config = Get-CurrentAudioConfig

    Write-Host ("`$DefaultPlaybackName       = `"$($config.DefaultPlaybackName)`"  # Multimedia Playback`n" +
           "`$DefaultCommPlaybackName   = `"$($config.DefaultCommPlaybackName)`"  # Communications Playback`n" +
           "`$DefaultRecordingName      = `"$($config.DefaultRecordingName)`"  # Multimedia Recording`n" +
           "`$DefaultCommRecordingName  = `"$($config.DefaultCommRecordingName)`"  # Communications Recording")

    exit
}

# Check if configuration file exists
if (-not (Test-Path $ConfigPath)) {
    Write-Error("`No configuration file found at $ConfigPath. Please create one by running `JankSwitch.ps1 -SaveConfig` after setting up your audio devices.")
    exit 1
}

# Load configuration from file
try {
    $ConfigContent = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json
    $DefaultPlaybackName = $ConfigContent.DefaultPlaybackName
    $DefaultCommPlaybackName = $ConfigContent.DefaultCommPlaybackName
    $DefaultRecordingName = $ConfigContent.DefaultRecordingName
    $DefaultCommRecordingName = $ConfigContent.DefaultCommRecordingName
}
catch {
    Write-Error ("Failed to load configuration from $ConfigPath. Please ensure the file is valid JSON.`nError: $($_.Exception.Message)")
    exit 1
}

# Get full list of playback and recording devices for reliable lookup
$allPlayback  = Get-AudioDevice -List | Where-Object { $_.Type -eq 'Playback' }
$allRecording = Get-AudioDevice -List | Where-Object { $_.Type -eq 'Recording' }

function Set-IfChanged {
    param(
        [ValidateSet('Playback','Recording')] [string]$Type,
        [ValidateSet('Multimedia','Communications')] [string]$Role,
        [string]$TargetName
    )

    # Get current device depending on role and type
    switch ("$Type-$Role") {
        "Playback-Multimedia" {
            $current = Get-AudioDevice -Playback | Where-Object { $_.Default -eq $true } | Select-Object -First 1
        }
        "Recording-Multimedia" {
            $current = Get-AudioDevice -Recording | Where-Object { $_.Default -eq $true } | Select-Object -First 1
        }
        "Playback-Communications" {
            $current = Get-AudioDevice -PlaybackCommunication | Where-Object { $_.DefaultCommunication -eq $true } | Select-Object -First 1
        }
        "Recording-Communications" {
            $current = Get-AudioDevice -RecordingCommunication | Where-Object { $_.DefaultCommunication -eq $true } | Select-Object -First 1
        }
        default {
            Write-Warning "Invalid Type-Role combination: $Type - $Role"
            return
        }
    }

    # Get the name of the current device
    $fromName = if ($current -and $current.Name) { $current.Name } else { "Unknown" }

    $allDevices = if ($Type -eq 'Playback') { Get-AudioDevice -List | Where-Object { $_.Type -eq 'Playback' } } else { Get-AudioDevice -List | Where-Object { $_.Type -eq 'Recording' } }
    $targetDevice = $allDevices | Where-Object { $_.Name -eq $TargetName }

    if (-not $targetDevice) {
        Write-Warning "Target device with name '$TargetName' not found for $Type $Role"
        return
    }

    if ($fromName -eq $TargetName) {
        return
    }

    if ($Role -eq 'Communications') {
        Set-AudioDevice -ID $targetDevice.ID -CommunicationOnly
    } else {
        Set-AudioDevice -ID $targetDevice.ID -DefaultOnly
    }

    Start-Sleep -Milliseconds 1000
}

Set-IfChanged -Type Playback  -Role Multimedia     -TargetName $DefaultPlaybackName
Set-IfChanged -Type Recording -Role Multimedia     -TargetName $DefaultRecordingName
Set-IfChanged -Type Playback  -Role Communications -TargetName $DefaultCommPlaybackName
Set-IfChanged -Type Recording -Role Communications -TargetName $DefaultCommRecordingName