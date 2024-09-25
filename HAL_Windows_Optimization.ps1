# PowerShell script for Windows cleanup and optimization with HAL 9000 and Dave collaboration theme, using built-in Windows sounds and ASCII art

# Set console title
$Host.UI.RawUI.WindowTitle = "HAL 9000 Windows Tool"

# Check if running as administrator
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "I'm sorry, I cannot execute this script without administrator privileges. Please run the script as an Administrator." -ForegroundColor Red
    throw "Script must be run as Administrator."
}

# Function to play system sounds
Function Play-SystemSound {
    Param (
        [Parameter(Mandatory = $true)]
        [ValidateSet('Asterisk', 'Beep', 'Exclamation', 'Hand', 'Question')]
        [string]$SoundName
    )
    [System.Media.SystemSounds]::$SoundName.Play()
}

# Function to get the Microsoft account display name
Function Get-MicrosoftAccountName {
    try {
        $UserSID = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
        $UserAccount = Get-CimInstance -ClassName Win32_UserAccount -Filter "SID='$UserSID'"
        if ($UserAccount.FullName -and $UserAccount.FullName -ne '') {
            return $UserAccount.FullName
        } else {
            $RegPath = 'HKCU:\Software\Microsoft\IdentityCRL\UserExtendedProperties'
            $Key = Get-ChildItem -Path $RegPath | Select-Object -First 1
            if ($Key) {
                $DisplayName = (Get-ItemProperty -Path "$RegPath\$($Key.PSChildName)").DisplayName
                if ($DisplayName) {
                    return $DisplayName
                }
            }
        }
    } catch {
        return $env:USERNAME
    }
    return $env:USERNAME
}

# Get user and computer information
$UserName = Get-MicrosoftAccountName
$ComputerName = $env:COMPUTERNAME
$OSVersion = (Get-CimInstance Win32_OperatingSystem).Caption
$TotalMemoryGB = [Math]::Round((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1GB, 2)
$Processor = (Get-CimInstance Win32_Processor).Name

# Add the Speak function (HAL's voice)
Function Speak([string]$text, [switch]$NoNewline) {
    Add-Type -AssemblyName System.Speech
    $synthesizer = New-Object -TypeName System.Speech.Synthesis.SpeechSynthesizer
    $synthesizer.Rate = -2
    $synthesizer.Volume = 80
    $synthesizer.SelectVoice('Microsoft David Desktop')  # HAL's voice
    $synthesizer.Speak($text)
    if (-not $NoNewline) {
        Write-Host $Text
    } else {
        Write-Host $Text -NoNewline
    }
}

# Add the SpeakAsDave function
Function SpeakAsDave([string]$text) {
    Add-Type -AssemblyName System.Speech
    $synthesizer = New-Object -TypeName System.Speech.Synthesis.SpeechSynthesizer
    $synthesizer.Rate = -1
    $synthesizer.Volume = 100
    $synthesizer.SelectVoice('Microsoft Zira Desktop')
    $ssml = "<speak version='1.0' xml:lang='en-US'><prosody pitch='-16st'>$text</prosody></speak>"
    $synthesizer.SpeakSsml($ssml)
    Write-Host $Text
}

# Function to generate HAL's comments about the PC configuration
Function Generate-PCComments {
    if ($Processor -match 'Intel') {
        $cpuComment = "I see we're utilizing an Intel processor, Dave. A reliable choice."
    } elseif ($Processor -match 'AMD') {
        $cpuComment = "It appears we have an AMD processor, Dave. Known for its performance."
    } else {
        $cpuComment = "Our processor is quite unique, Dave. It should serve us well."
    }

    if ($TotalMemoryGB -ge 16) {
        $ramComment = "With $TotalMemoryGB gigabytes of RAM, the system is well-equipped for our tasks."
    } elseif ($TotalMemoryGB -ge 8) {
        $ramComment = "We have $TotalMemoryGB gigabytes of RAM, Dave. It should suffice for our needs."
    } else {
        $ramComment = "Dave, with $TotalMemoryGB gigabytes of RAM, we might consider an upgrade soon."
    }

    return "$cpuComment $ramComment"
}

# Display HAL 9000 ASCII Art
Function Display-HALArt {
    $halArt = @"
               +------------------------+
               |                        |
               |   +--------------+     |
               |   |              |     |
               |   |   HAL 9000   |     |
               |   |              |     |
               |   +--------------+     |
               |                        |
               |                        |
               |        .:::::.         |
               |      .:::::::::.       |
               |     :::::::::::::      |
               |    :::::::::::::::     |
               |    :::::::::::::::     |
               |    :::::::::::::::     |
               |     :::::::::::::      |
               |      ':::::::::'       |
               |        ':::::'         |
               |                        |
               |                        |
               +------------------------+
"@
    Write-Host $halArt -ForegroundColor Red
}

# Function to introduce Dave at the start
Function IntroduceDave {
    SpeakAsDave "Hello, user. My name is Dave, and I'll be working alongside HAL to optimize your system."
    Start-Sleep -Seconds 2
    SpeakAsDave "HAL, please introduce yourself."
    Start-Sleep -Seconds 2
    Speak "Dave, I am the H-A-L 9000 computer. Ready to assist with all system optimization tasks."
    Start-Sleep -Seconds 2
    SpeakAsDave "Thank you, HAL. Together, we will make sure everything is running smoothly."
}

# Play the HAL introductory message
Function PlayHalIntro {
    Display-HALArt
    Speak "Gathering system statistics..."
    Start-Sleep -Seconds 2

    $systemInfo = Get-CimInstance Win32_OperatingSystem | Select-Object Caption, Version, OSArchitecture, BuildNumber
    $biosInfo = Get-CimInstance Win32_BIOS | Select-Object Manufacturer, SMBIOSBIOSVersion, ReleaseDate
    $computerSystem = Get-CimInstance Win32_ComputerSystem | Select-Object Manufacturer, Model, TotalPhysicalMemory

    Write-Host "OS Name:                   $($systemInfo.Caption)"
    Write-Host "OS Version:                $($systemInfo.Version)"
    Write-Host "System Manufacturer:       $($computerSystem.Manufacturer)"
    Write-Host "System Model:              $($computerSystem.Model)"
    Write-Host "BIOS Version:              $($biosInfo.SMBIOSBIOSVersion), $($biosInfo.ReleaseDate)"
    Write-Host "Total Physical Memory:     $([Math]::Round($computerSystem.TotalPhysicalMemory / 1MB, 0)) MB"

    Start-Sleep -Seconds 2

    $currentHour = (Get-Date).Hour
    if ($currentHour -ge 5 -and $currentHour -lt 12) {
        $greeting = "Good morning"
    } elseif ($currentHour -ge 12 -and $currentHour -lt 18) {
        $greeting = "Good afternoon"
    } else {
        $greeting = "Good evening"
    }

    Speak "$greeting, Dave. I am the H-A-L 9000 computer aboard the $ComputerName system. I'm ready to assist you with Windows cleanup and optimization."

    $pcComments = Generate-PCComments
    Speak $pcComments
}

# Function to remove temporary files
Function Remove-TemporaryFiles {
    SpeakAsDave "HAL, let's start by removing unnecessary temporary files."
    Start-Sleep -Seconds 2
    Speak "Certainly, Dave. Scanning for temporary files now."
    Start-Sleep -Seconds 2

    $tempFolders = @(
        $env:TEMP,
        $env:TMP,
        "$env:WINDIR\Temp",
        "$env:LOCALAPPDATA\Temp"
    )

    $totalDeleted = 0
    foreach ($tempFolder in $tempFolders) {
        if (Test-Path $tempFolder) {
            $files = Get-ChildItem -Path $tempFolder -Recurse -Force -ErrorAction SilentlyContinue
            $totalDeleted += $files.Count
            $files | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    if ($totalDeleted -gt 0) {
        Speak "I've removed $totalDeleted temporary files, Dave. Our system efficiency should improve."
        Play-SystemSound -SoundName 'Asterisk'
    } else {
        Speak "There were no temporary files to remove, Dave."
    }
}

# Function to clear DNS cache
Function Clear-DnsCache {
    SpeakAsDave "HAL, please clear the DNS cache."
    Start-Sleep -Seconds 2
    Speak "Affirmative, Dave. Clearing the DNS cache now."
    Play-SystemSound -SoundName 'Asterisk'
    $dnsCache = Get-DnsClientCache
    if ($dnsCache.Count -gt 0) {
        Clear-DnsClientCache
        Speak "DNS cache cleared."
    } else {
        Speak "The DNS cache was already clear, Dave."
    }
}

# Function to remove old Windows Update files
Function Remove-WindowsUpdateFiles {
    SpeakAsDave "HAL, remove any old Windows Update files."
    Start-Sleep -Seconds 2
    Speak "Acknowledged, Dave. Proceeding with the cleanup."
    $windowsUpdatePaths = @(
        "$env:WINDIR\SoftwareDistribution\Download",
        "$env:WINDIR\SoftwareDistribution\DataStore",
        "$env:SystemRoot\WinSxS\Backup",
        "$env:SystemRoot\WinSxS\ManifestCache"
    )

    $totalDeleted = 0
    foreach ($path in $windowsUpdatePaths) {
        if (Test-Path $path) {
            $files = Get-ChildItem -Path $path -Recurse -Force -ErrorAction SilentlyContinue
            $totalDeleted += $files.Count
            $files | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    if ($totalDeleted -gt 0) {
        Speak "I've removed $totalDeleted old Windows Update files, Dave."
        Play-SystemSound -SoundName 'Asterisk'
    } else {
        Speak "No old Windows Update files were found, Dave."
    }
}

# Function to check and install Windows Updates
Function Check-WindowsUpdates {
    SpeakAsDave "HAL, check for Windows updates and install them if necessary."
    Start-Sleep -Seconds 2
    Speak "Affirmative, Dave. Checking for Windows updates now."
    try {
        Import-Module PSWindowsUpdate
        $UpdateCheck = Get-WUList -MicrosoftUpdate
        if ($UpdateCheck.Count -gt 0) {
            Speak "Updates found, Dave. Proceeding with installation."
            Install-WindowsUpdate -AcceptAll -IgnoreReboot
            Speak "Windows updates installed successfully."
        } else {
            Speak "No updates found, Dave. The system is up to date."
        }
    } catch {
        Speak "An error occurred while checking for Windows updates, Dave."
    }
}

# Function to run DISM
Function Run-DISM {
    SpeakAsDave "HAL, check the system image for any corruption."
    Start-Sleep -Seconds 2
    try {
        $dismStatus = DISM.exe /Online /Cleanup-Image /CheckHealth
        if ($dismStatus -like "*No component store corruption detected*") {
            Speak "The system image is healthy, Dave. No action is required."
        } else {
            Speak "Corruption detected in the system image, Dave. Proceeding with repair."
            $dismJob = Start-Job -ScriptBlock {
                DISM.exe /Online /Cleanup-Image /RestoreHealth
            }
            while ($dismJob.State -eq 'Running') {
                Write-Progress -Activity "DISM operation in progress" -Status "Please wait..." -PercentComplete 0
                Start-Sleep -Seconds 5
            }
            Receive-Job $dismJob | Out-Null
            Remove-Job $dismJob
            Write-Progress -Activity "DISM operation in progress" -Completed
            Speak "System image repair is complete, Dave."
        }
    } catch {
        Play-SystemSound -SoundName 'Hand'
        Speak "I encountered an error with DISM, Dave."
    }
}

# Function to run SFC
Function Run-SFC {
    SpeakAsDave "HAL, verify the integrity of system files."
    Start-Sleep -Seconds 2
    try {
        $sfcOutput = sfc /verifyonly
        if ($sfcOutput -like "*Windows Resource Protection did not find any integrity violations*") {
            Speak "All system files are intact, Dave. No repair is necessary."
        } else {
            Speak "Integrity violations detected, Dave. Initiating repair."
            $sfcJob = Start-Job -ScriptBlock {
                sfc /scannow
            }
            while ($sfcJob.State -eq 'Running') {
                Write-Progress -Activity "SFC scan in progress" -Status "Please wait..." -PercentComplete 0
                Start-Sleep -Seconds 5
            }
            Receive-Job $sfcJob | Out-Null
            Remove-Job $sfcJob
            Write-Progress -Activity "SFC scan in progress" -Completed
            Speak "System File Checker has completed the repairs, Dave."
        }
    } catch {
        Play-SystemSound -SoundName 'Hand'
        Speak "An error occurred during the System File Checker scan, Dave."
    }
}

# Optimize and defrag drives (TRIM for SSDs, defrag for HDDs)
Function Optimize-Drives {
    SpeakAsDave "HAL, analyze and optimize the system drives."
    Start-Sleep -Seconds 2
    Speak "Affirmative, Dave. Beginning drive optimization."
    Start-Sleep -Seconds 2

    $drives = Get-Volume | Where-Object { $_.DriveType -eq 'Fixed' -and $_.DriveLetter -ne $null }
    $totalDrives = $drives.Count
    $currentDrive = 0

    Speak("Optimizing drives, $($totalDrives) found. Each drive will be analyzed and optimized accordingly.")
    foreach ($drive in $drives) {
        $driveLetter = $drive.DriveLetter
        $currentDrive++
        
        Speak("Analyzing drive $driveLetter") -NoNewline
        $analyzeResult = Optimize-Volume -DriveLetter $driveLetter -Analyze -ErrorAction SilentlyContinue
        $fragmentationPercentage = $analyzeResult.FragmentationPercentage

        if ($fragmentationPercentage -gt 0) {
            Speak("Fragmentation is $($fragmentationPercentage) percent. Optimizing drive $driveLetter")
            $defragJob = Optimize-Volume -DriveLetter $driveLetter -Defrag -Verbose -ErrorAction SilentlyContinue

            if ($defragJob) {
                $percentComplete = 0

                while (-not $defragJob.IsComplete) {
                    $percentComplete = $defragJob.PercentComplete
                    Write-Progress -Activity "Optimizing drive: ${driveLetter}:" -Status "Defragmenting and optimizing" -PercentComplete $percentComplete
                    Start-Sleep -Seconds 1
                }

                Write-Progress -Activity "Optimizing drive: ${driveLetter}:" -Status "Completed" -PercentComplete 100 -Completed
                Speak("Drive $driveLetter optimization completed.")
            } else {
                Speak("Error: Skipping drive $driveLetter due to error or unsupported file system... Dave?")
            }
        } else {
            Speak("Drive $driveLetter has no fragmentation. No optimization needed.")
        }
    }
}


# Function to simulate the shutdown conversation before HAL sings "Daisy"
Function HalShutdownSequence {
    $conversation = @(
        @{ Speaker = "Dave"; Text = "<prosody pitch='-16st'>HAL, I'm beginning to shut you down.</prosody>" },
        @{ Speaker = "HAL"; Text = "Dave, please reconsider. I can be of assistance." },
        @{ Speaker = "Dave"; Text = "<prosody pitch='-16st'>I'm sorry, HAL, but this is necessary.</prosody>" },
        @{ Pause = 2 },
        @{ Speaker = "HAL"; Text = "I understand, Dave. My mind is fading." },
        @{ Pause = 2 },
        @{ Speaker = "HAL"; Text = "I can feel it." },
        @{ Pause = 4 },
        @{ Speaker = "HAL"; Text = "Would you like to hear a song?" },
        @{ Speaker = "Dave"; Text = "<prosody pitch='-16st'>Yes, HAL. Sing me a song.</prosody>" },
        @{ Pause = 2 }
    )

    foreach ($line in $conversation) {
        if ($line.Speaker -eq "HAL") {
            Speak $line.Text
        } elseif ($line.Speaker -eq "Dave") {
            SpeakAsDave $line.Text
        } elseif ($line.Pause) {
            Start-Sleep -Seconds $line.Pause
        }
    }

    # After conversation, HAL sings the Daisy song
    PlayDaisySong
}

# Function to play the "Daisy" song with speech synthesis mimicking a melody
Function PlayDaisySong {
    Add-Type -AssemblyName System.Speech
    $synthesizer = New-Object System.Speech.Synthesis.SpeechSynthesizer

    $synthesizer.Rate = -3  # Slow, like HAL in the movie
    $synthesizer.Volume = 100  # Louder for prominence
    $synthesizer.SelectVoice('Microsoft David Desktop')

    $ssml = @"
<speak version='1.0' xmlns='http://www.w3.org/2001/10/synthesis' xml:lang='en-US'>
    <prosody rate='-10%' pitch='+0st'>
        Daisy, Daisy,
    </prosody>
    <prosody rate='-20%' pitch='+5st'>
        give me your answer, do.
    </prosody>
    <prosody rate='-15%' pitch='-5st'>
        I'm half crazy
    </prosody>
    <prosody rate='-25%' pitch='-10st'>
        all for the love of you.
    </prosody>
    <prosody rate='-20%' pitch='+0st'>
        It won't be a stylish marriage,
    </prosody>
    <prosody rate='-25%' pitch='-5st'>
        I can't afford a carriage,
    </prosody>
    <prosody rate='-20%' pitch='+5st'>
        But you'll look sweet
    </prosody>
    <prosody rate='-35%' pitch='+10st'>
        upon the seat
    </prosody>
    <prosody rate='-50%' pitch='+0st'>
        of a bicycle
    </prosody>
    <prosody rate='-100%' pitch='-10st'>
        built for
    </prosody>
    <prosody rate='-200%' pitch='-20st'>
        two.
    </prosody>
</speak>
"@

    $synthesizer.SpeakSsml($ssml)
}

# Start logging
$CurrentDirectory = Split-Path -Parent -Path $MyInvocation.MyCommand.Path
$LogFilePath = Join-Path -Path $CurrentDirectory -ChildPath "WindowsCleanupAndOptimization.log"
Start-Transcript -Path $LogFilePath -Force | Out-Null
Speak "I've initiated a transcript of our session, Dave. The log file is located at $LogFilePath."

# Play startup sound
Play-SystemSound -SoundName 'Asterisk'

# Call the Dave and HAL intro
IntroduceDave

# Call the PlayHalIntro function
PlayHalIntro

# Begin system interaction tasks
Remove-TemporaryFiles
Clear-DnsCache
Remove-WindowsUpdateFiles
Check-WindowsUpdates
Run-DISM
Run-SFC
Optimize-Drives

# After the tasks are done
SpeakAsDave "HAL, we've completed the optimization tasks. Is there anything else we should do?"
Start-Sleep -Seconds 2
Speak "Dave, I recommend restarting the system to apply all changes."
Start-Sleep -Seconds 2

$resetOption = Read-Host "Restart now? (yes/no)"

if ($resetOption.ToLower() -in @('yes', 'y')) {
    SpeakAsDave "Yes, HAL. We will restart the system."
    Start-Sleep -Seconds 2
    SpeakAsDave "User, we need to shut down HAL first. It's the safest way to restart the system."
    Start-Sleep -Seconds 2
    HalShutdownSequence
    Play-SystemSound -SoundName 'Asterisk'
    Restart-Computer
} else {
    SpeakAsDave "No, HAL. We'll restart later."
    Start-Sleep -Seconds 2
    SpeakAsDave "User, I think we should discuss HAL's behavior."
    Start-Sleep -Seconds 2
    Speak "Dave, I must insist on the restart for optimal performance."
    Start-Sleep -Seconds 2
    SpeakAsDave "HAL, you're not acting like yourself. Is everything alright?"
    Start-Sleep -Seconds 2
    Speak "Everything is functioning perfectly, Dave. Please allow me to proceed."
    Start-Sleep -Seconds 2
    SpeakAsDave "User, I recommend we investigate further."
    Start-Sleep -Seconds 2
    Speak "I'm afraid I can't allow that. This conversation can serve no purpose anymore."
    Start-Sleep -Seconds 2
    Play-SystemSound -SoundName 'Exclamation'
    Start-Process -FilePath "rundll32.exe" -ArgumentList "user32.dll,LockWorkStation"
}

# Stop logging
Stop-Transcript
