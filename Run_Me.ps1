##############################################
##        Run Me First                      ##
##############################################
$R1 = "1"
$Global:ErrorActionPreference = 'Stop'
$Global:Date = Get-Date -Format "dd-MM-yyyy_HH_mm"
$Global:E0 =  $PSScriptRoot 
$Global:Es =  $PSScriptRoot  +"\" + "Main Script.psm1"
$RO = $PSScriptRoot  +"\" + "RunOnce.psm1"
$Global:Ma32 = "C:\Temp\Scripts"
$script:E1 = "\"
$Script:E2 =  "Main Script.psm1"
$Global:E3 = $Ma32 + $E1 + $E2
$Global:E4 = $Ma32 + "\Errors\"
$Global:Hu = "100"
$Script:ERL4 = $E4 + "Run_Errors_$Date.txt"
$Script:ECL4 = $E4 + "Run_Exceptions_$Date.txt"
Try{
Function Check{
if( -Not (Test-Path -Path $Ma32 ) )

{

    New-Item -ItemType directory -Path $Ma32 |out-null

}
}
Function RunOnce{
    Import-Module $RO
}
Function Run{
copy-item $Es $Ma32
Import-Module $E3 -force
}

Function Global:ER{
if( -Not (Test-Path -Path $E4 ) )

{

    New-Item -ItemType directory -Path $E4 |out-null

}
}

Function St{
If($R1 -le "1"){
#RunOnce
}
Check
Run
}
St
}
Catch {
ER
$RMError = $_ 
$RMError2 = $_.exception.Message
$RMError| Out-File $ERL4 -append
$RMError2| Out-File $ECL4 -append
$Res = "Script Ended Prematurely in Run Me please see $ERL4"
$a = new-object -comobject wscript.shell
$b = $a.popup("$Res ",3," Error ",48)
Write-Host "Script Ended Prematurely in Run Me please see $ERL4" -ForegroundColor Magenta
Return
}