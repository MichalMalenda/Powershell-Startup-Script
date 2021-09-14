###
#PLEASE ENSURE THAT SYNTH IS ALWAYS ON THE FIRST POSITION IN THE SETTINGS
#
#TO ENABLE OR DISABLE ONE OF THE STARTUP APPS YOU ARE FREE TO HANGE THE DECISION FROM TRUE TO FALSE
#
#PLEASE ENSURE THAT THE ACTUAL PROCESS NAME OF THE RUNNING APP IS IN THE POSITION OF A NAME UNDER SETTINGS
#
#PATH=[SETTING] MEANS THAT THE VALUE IS THERE TO BE SET AND NOT ACTIONED
#
#PATH=[FUNCTION] MEANS THAT THE VALUE IS THERE TO BE ACTIONED AS A FUNCTION
#
#PATH=[ACTUAL PATH] FOR THE APPLICATION MEANS THAT IT WILL BE INVOKED AS AN .EXE FILE
#
###
$ErrorActionPreference = “SilentlyContinue”
$GLOBAL:SETTINGS=@(
    [PSCustomObject]@{NAME='SYNTH';DECISION='ON';PATH='SETTING'}
    [PSCustomObject]@{NAME='Disallow-Shaking';DECISION='TRUE';PATH='FUNCTION'}
    [PSCustomObject]@{NAME='Office-Theme-Set';DECISION='TRUE';PATH='FUNCTION'}
    [PSCustomObject]@{NAME='Set-Wallpaper';DECISION='FALSE';PATH='FUNCTION'}
    [PSCustomObject]@{NAME='Enable-Nav-Pen';DECISION='FALSE';PATH='FUNCTION'}
    [PSCustomObject]@{NAME='OUTLOOK';DECISION='TRUE';PATH='C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE'}
    [PSCustomObject]@{NAME='MSTSC';DECISION='TRUE';PATH='C:\Windows\System32\mstsc.exe'}#RDP
    [PSCustomObject]@{NAME='ONENOTE';DECISION='TRUE';PATH='C:\Program Files\Microsoft Office\root\Office16\ONENOTE.EXE'}
    [PSCustomObject]@{NAME='PTONECLK';DECISION='FALSE';PATH='C:\Program Files (x86)\Webex\Webex\Applications\ptoneclk.exe'}#Webex
    [PSCustomObject]@{NAME='CODE';DECISION='FALSE';PATH='C:\Program Files\Microsoft VS Code\Code.exe'}#VSCode
    [PSCustomObject]@{NAME='BROWSER';DECISION='TRUE';PATH='C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'}
)
if($SETTINGS[0].NAME -like "SYNTH" -and $SETTINGS[0].DECISION -like "ON"){
    Add-Type -AssemblyName System.Speech
    $Synth = New-Object -TypeName System.Speech.Synthesis.SpeechSynthesizer
}else{
    $Synth = $false
}
Function Disallow-Shaking{
    $check_if_exists=Get-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "DisallowShaking"
    if(!$check_if_exists){
        New-ItemProperty –Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" –Name "DisallowShaking" -Value 1 -PropertyType "DWord"
        Write-Output "-> Disallow Shaking had to be enabled again"
        $Synth.Speak("Disallow Shaking had to be enabled again.")
    } else{
        Write-Output "-> There was no need to disallow Shaking"
        $Synth.Speak("There was no need to disallow Shaking.")
    }
}
Function Office-Theme-Set{
    $check_office_theme=(Get-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Common").'UI Theme'
    if($check_office_theme -ne "4"){
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Common" -Name "UI Theme" -Value 4 -Type DWORD
        Get-ChildItem -Path ("HKCU:\Software\Microsoft\Office\16.0\Common" + "\Roaming\Identities\") | ForEach-Object {
            $identityPath=($_.Name.Replace("HKEY_CURRENT_USER", "HKCU:") + "\Settings\1186\{00000000-0000-0000-0000-000000000000}");
            if(Get-ItemProperty -Path $identityPath -Name "Data" -ErrorAction Ignore){
            Set-ItemProperty -Path $identityPath -Name "Data" -Value ([byte[]](4, 0, 0, 0)) -Type Binary
            }
        }
        Write-Output "-> Office theme had to be changed on startup"
        $Synth.Speak("Office theme had to be changed on startup.")
    }else{
        Write-Output "-> There was no need to change office theme"
        $Synth.Speak("There was no need to change office theme.")
    }
}
Function Enable-Nav-Pen{

}
Function Set-Wallpaper{
    $out_file_name="C:\temp\wallpaper1920x1300.jpg"
    $url = "https://images.pexels.com/photos/3586966/pexels-photo-3586966.jpeg?auto=compress&cs=tinysrgb&dpr=3&h=750&w=1260"
    $wc = New-Object System.Net.WebClient
    $wc.DownloadFile($url, $out_file_name)
    Add-Type -AssemblyName System.Windows.Forms
    $Monitors = [System.Windows.Forms.Screen]::AllScreens
    Set-ItemProperty -path 'HKCU:\Control Panel\Desktop\' -name Wallpaper -value $out_file_name
    rundll32.exe user32.dll, UpdatePerUserSystemParameters
    rundll32.exe user32.dll, UpdatePerUserSystemParameters
    Start-Sleep 2
    Remove-Item -Path $out_file_name -Force
}
Function Check-n-run{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Process_Name,
        [Parameter(Mandatory)]
        [string]$Path
    )
    $temp=Get-Process -Name $Process_Name
    if(!$temp){
        Verify-Invoke -Path $Path | Out-Null
    }else{
        Write-Output "-> $Process_Name is already running"
    }
}
Function Verify-Invoke{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [Parameter()]
        [string]$Website
    )
    begin{
        if($Path -and !$Website){
            $pick="APP"
        }else{
            $pick="WEBSITE"
        }
    }
    process{
        switch($pick){
            "APP"{
                $result=Test-Path -Path $Path
                if(!$result){
                    $message="-> ERROR: Cannot open $Path as it is not found"
                }else{
                    $message="-> SUCCESS: Opening $Path"
                    Invoke-Item -Path $Path
                }
            }
            "WEBSITE"{
                $result=Test-Path -Path $Path
                if(!$result){
                    $message="-> ERROR: Cannot open $Path as it is not found"
                }else{
                    $message="-> SUCCESS: Opening $Website"
                    Start-Process $Path $Website
                }
            }
        }
    }
    end{
        Write-Output $message
    }
}
Foreach($option in $SETTINGS){
    if($option.DECISION -like "TRUE"){
        if($option.PATH -like "FUNCTION"){
            Invoke-Expression $option.NAME
        }elseif($option.NAME -notlike "BROWSER"){
            Check-n-run -Process_Name $option.NAME -Path $option.PATH
        }else{
            Verify-Invoke -Path $option.PATH -Website "https://google.com" | Out-Null
        }
    }
}