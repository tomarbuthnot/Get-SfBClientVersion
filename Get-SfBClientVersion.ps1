# Currently works for Office 2016 x86 in the default program files path

$SFBEXE = Get-Item ("${Env:ProgramFiles(x86)}" + "\Microsoft Office\root\Office16\lync.exe")
$SfBMSO = Get-Item ("${Env:ProgramFiles(x86)}" + "\Microsoft Office\root\Office16\MSOSB.DLL")

$sfbexeVersion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($SFBEXE)
$sfbmsoVersion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($SfBMSO)

Write-Host " "
Write-Host "SFB EXE Version: $($sfbexeVersion.fileversion)"
Write-Host "SFB MSO Version: $($sfbmsoVersion.fileversion)"

<#
# If Sigcheck exists check timestamps

$SigCheck = "C:\Dropbox\PowerShell\SysInternals\sigcheck.exe"
# Windows SysInternals Sigcheck https://technet.microsoft.com/en-gb/sysinternals/bb897441.aspx


IF ((Test-Path $SigCheck) -eq $True)
    {
    Write-Host "SigCheck Found"
    $SfBexeSigcheck = Start-Process -FilePath $SigCheck -ArgumentList "'C:\Program Files (x86)\Microsoft Office\root\Office16\lync.exe'" -WindowStyle Normal | Out-String

    }
#>
