$Script:ERL7 = $E4+ "Again_Errors_$Date.txt"
$Script:ECL7 = $E4 + "Again_Exceptions_$Date.txt"
Function Get-Choice{
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

$Form1 = New-Object System.Windows.Forms.Form

$Form1.Text = "Second Choice"

$Form1.ClientSize = New-Object System.Drawing.Size(250, 90)

$form1.topmost = $true

$buttonExit_Click={
 $cancel = $True
 $Form1.Close()
 }

$Text = New-Object System.Windows.Forms.Label

$Text.Location = New-Object System.Drawing.Point(15, 15)

$Text.Size = New-Object System.Drawing.Size(250, 20)

$Text.Text = "Would you like to run the script again?"

$Form1.Controls.Add($Text)

#$ErrorActionPreference = "SilentlyContinue"

Function Button1

{

$Button1 = New-Object System.Windows.Forms.Button

$Button1.Location = New-Object System.Drawing.Point(15, 50)

$Button1.Size = New-Object System.Drawing.Size(100, 20)

$Button1.Text = "Yes"

$Button1.add_Click({Get-Script

$Form1.Close()})

$Form1.Controls.Add($Button1)

}

Function Button2

{

$Button2 = New-Object System.Windows.Forms.Button

$Button2.Location = New-Object System.Drawing.Point(125, 50)

$Button2.Size = New-Object System.Drawing.Size(100, 20)

$Button2.Text = "No"

$Button2.add_Click({Get-Old

$Form1.Close()})

$Form1.Controls.Add($Button2)

}

Button1

Button2

[void]$form1.showdialog()
}
Function Get-Script{
$Form1.Dispose()
Import-Module $E3 -force

}
Function Get-old{
Return
}
Try{
Get-Choice
}
Catch{
ER
$AGError2 = $_.exception.Message
$AGError = $_ 
$AGError2| Out-File $ECL7 -append
$AGError | Out-File $ERL7
Write-Host "Script Ended Prematurely please see $ERL7" -ForegroundColor Magenta
Return
}