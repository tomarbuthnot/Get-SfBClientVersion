# Currently works for Office 2016 x86 in the default program files path


# Office 2016 x86 C2R

$2016CR2x86Test = Test-Path ("${Env:ProgramFiles(x86)}" + "\Microsoft Office\root\Office16\lync.exe")

IF ($2016CR2x86Test -eq $true)
{
$SFBEXE = Get-Item ("${Env:ProgramFiles(x86)}" + "\Microsoft Office\root\Office16\lync.exe")
$SfBMSO = Get-Item ("${Env:ProgramFiles(x86)}" + "\Microsoft Office\root\Office16\MSOSB.DLL")
$InstallType = "Office 2016 x86 C2R"
}

# Office 2013 x86 MSI


$2016MSIx86Test = Test-Path ("${Env:ProgramFiles(x86)}" + "\Microsoft Office\Office16\lync.exe")

IF ($2016MSIx86Test -eq $true)
{
$SFBEXE = Get-Item ("${Env:ProgramFiles(x86)}" + "\Microsoft Office\Office16\lync.exe")
$SfBMSO = Get-Item ("${Env:ProgramFiles(x86)}" + "\Microsoft Office\Office16\MSOSB.DLL")
$InstallType = "Office 2016 x86 MSI"
}






###################### Office 2013 ################################


# Office 2013 x86 C2R

$2013CR2x86Test = Test-Path ("${Env:ProgramFiles(x86)}" + "\Microsoft Office 15\root\office15\lync.exe")

IF ($2013CR2x86Test -eq $true)
{
$SFBEXE = Get-Item ("${Env:ProgramFiles(x86)}" + "\Microsoft Office 15\root\office15\lync.exe")
$SfBMSO = Get-Item ("${Env:ProgramFiles(x86)}" + "\Microsoft Office 15\root\office15\MSOSB.DLL")
$InstallType = "Office 2013 x86 C2R"
}


# Office 2013 x64 C2R

$2013CR2x64Test = Test-Path ("${Env:ProgramFiles}" + "\Microsoft Office 15\root\office15\lync.exe")

IF ($2013CR2x64Test -eq $true)
{
$SFBEXE = Get-Item ("${Env:ProgramFiles}" + "\Microsoft Office 15\root\office15\lync.exe")
$SfBMSO = Get-Item ("${Env:ProgramFiles}" + "\Microsoft Office 15\root\office15\MSOSB.DLL")
$InstallType = "Office 2013 x64 C2R"
}



# Office 2013 x86 MSI

# $2013MSIx86Test = Test-Path ("${Env:ProgramFiles(x86)}" + "\Microsoft Office\Office16\lync.exe")

IF ($2013MSIx86Test -eq $true)
{
$SFBEXE = Get-Item ("${Env:ProgramFiles(x86)}" + "\Microsoft Office\Office16\lync.exe")
$SfBMSO = Get-Item ("${Env:ProgramFiles(x86)}" + "\Microsoft Office\Office16\MSOSB.DLL")
$InstallType = "Office 2016 x86 MSI"
}



# Output

$sfbexeVersion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($SFBEXE)
$sfbmsoVersion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($SfBMSO)

Write-Host " "
Write-Host "$InstallType"
Write-Host "SFB EXE Version: $($sfbexeVersion.fileversion)"
Write-Host "SFB MSO Version: $($sfbmsoVersion.fileversion)"
Write-Host " "

