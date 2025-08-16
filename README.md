# JankSwitch
A janky PowerShell script for managing Windows default audio devices via a config file

## How It Works
This script saves and restores your preferred audio device settings for:
- Default playback device (speakers/headphones)
- Communications playback device
- Default recording device (microphone)
- Communications recording device

## Usage
- Get current config: `JankSwitch.ps1 -ListConfig`
- Save current config: `JankSwitch.ps1 -SaveConfig`
- Apply saved config: Just run `JankSwitch.exe`

## My Setup
I packaged the PowerShell script to an EXE using:
```
Invoke-PS2EXE .\JankSwitch.ps1 .\JankSwitch.exe -noConsole
```

Using Windows Task Scheduler, I configured the executable to run:
- At system startup
- Every 5 minutes as a recurring task

I have included `audio.conf.example` which is my personal config that sets Voicemeeter as my default audio device.

My audio settings automatically revert to my defaults whenever Nvidia decides to change them.

## Requirements
- AudioDeviceCmdlets
- PS2EXE (only needed if building the EXE) 