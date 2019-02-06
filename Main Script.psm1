##############################################
##        Main Script                       ##
##############################################
Write-host $E0
$Script:Ma33 = "C:\Temp\Scripts\"
$M = "\Members_Complete.psm1"
$F = "\Folder_Permissions_Complete.psm1"
$S = "\Shares_Complete.psm1"
$G = "\Again.psm1"
$A = "\Auto_Folders.psm1"
$Script:Sc1 = $E0 + "\*.Psm1"
$Script:Ns1 = $PSScriptRoot + $F
$Script:Ns2 = $PSScriptRoot + $M
$Script:Ns3 = $PSScriptRoot + $S
$Global:Ns4 = $PSScriptRoot + $G
$Script:Ns5 = $PSScriptRoot + $A
$Global:terminateScript = $false
$Script:ERL5 = $E4+ "Main_Errors_$Date.txt"
$Script:ECL5 = $E4 + "Main_Exceptions_$Date.txt"
Try{
Function Loading {
        
    $Res = "Script is loading"
    $Res2 = "Please Wait.."
    $a = new-object -comobject wscript.shell
    $b = $a.popup("$Res`n$Res2 ",1," Information ",65)
    $Res2 = "Please Wait...."
    $a = new-object -comobject wscript.shell
    $b = $a.popup("$Res`n$Res2 ",1," Information ",65)
    $Res2 = "Please Wait......"
    $a = new-object -comobject wscript.shell
    $b = $a.popup("$Res`n$Res2 ",1," Information ",65)
    $Res2 = "Please Wait........"
    $a = new-object -comobject wscript.shell
    $b = $a.popup("$Res`n$Res2 ",1," Information ",65)
    $Res2 = "Please Wait.........."
    $a = new-object -comobject wscript.shell
    $b = $a.popup("$Res`n$Res2 ",1," Information ",65)
}
Function Get-Scripts{
    Loading
copy-item $Sc1 $Ma33
}
Function Global:StartProgressBar1{
    Try{
	if($i -le 5){
        $Global:i += 1
	}
	else {
        $timer.enabled = $false 
    }
}
Catch [System.Management.Automation.RuntimeException]{
    ER
    $MSError =$_
    $MSError2 = $_.exception.Message
    $MSError| Out-File $ERL5 -append
    $MSError2 | Out-File $ECL5 -append
    Write-Host "Script Ended Prematurely in Main Script please see $ERL5" -ForegroundColor Magenta
    Return
}
}
Function Global:ProgressBar1{
    Try{
    Add-Type -assembly System.Windows.Forms
    ## -- Create The Progress-Bar
    $Global:ObjForm = New-Object System.Windows.Forms.Form
    $ObjForm.Text = "Searching for shares on Servers"
    $ObjForm.Height = 100
    $ObjForm.Width = 500
    $ObjForm.BackColor = "White"

    $ObjForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
    $ObjForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

    ## -- Create The Label
    $ObjLabel = New-Object System.Windows.Forms.Label
    $ObjLabel.Text = "Starting. Please wait ... "
    $ObjLabel.Left = 5
    $ObjLabel.Top = 10
    $ObjLabel.Width = 500 - 20
    $ObjLabel.Height = 15
    $ObjLabel.Font = "Tahoma"
    ## -- Add the label to the Form
    $ObjForm.Controls.Add($ObjLabel)

    $Global:PB = New-Object System.Windows.Forms.ProgressBar
    $PB.Name = "PowerShellProgressBar"
    $PB.Value = 0
    $PB.Style="Continuous"

    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 500 - 40
    $System_Drawing_Size.Height = 20
    $PB.Size = $System_Drawing_Size
    $PB.Left = 5
    $PB.Top = 40
    $ObjForm.Controls.Add($PB)
##############################################################
    $i = 0
    $Global:timer = New-Object System.Windows.Forms.Timer 
    $timer.Interval = 1000

    $timer.add_Tick({
        
        StartProgressBar1
        
    })

    $timer.Enabled = $true
    $timer.Start()
##############################################################
    ## -- Show the Progress-Bar and Start The PowerShell Script
    $ObjForm.Show() | Out-Null
    $ObjForm.Focus() | Out-NUll
    $ObjLabel.Text = "Starting. Please wait ... "
    $ObjForm.Refresh()

    Start-Sleep -Seconds 1
    Out-Null
}
Catch [System.Management.Automation.RuntimeException]{
    ER
    $MSError =$_
    $MSError2 = $_.exception.Message
    $MSError| Out-File $ERL5 -append
    $MSError2 | Out-File $ECL5 -append
    Write-Host "Script Ended Prematurely in Main Script please see $ERL5" -ForegroundColor Magenta
    Return
}
}

Function Global:ProgressBar2{
    $ErrorActionPreference = 'Stop'
    Add-Type -assembly System.Windows.Forms
    ## -- Create The Progress-Bar
    $Global:ObjForm = New-Object System.Windows.Forms.Form
    $ObjForm.Text = "Creating xlsx File"
    $ObjForm.Height = 100
    $ObjForm.Width = 500
    $ObjForm.BackColor = "White"

    $ObjForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
    $ObjForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

    ## -- Create The Label
    $ObjLabel = New-Object System.Windows.Forms.Label
    $ObjLabel.Text = "Starting. Please wait ... "
    $ObjLabel.Left = 5
    $ObjLabel.Top = 10
    $ObjLabel.Width = 500 - 20
    $ObjLabel.Height = 15
    $ObjLabel.Font = "Tahoma"
    ## -- Add the label to the Form
    $ObjForm.Controls.Add($ObjLabel)

    $Global:PB = New-Object System.Windows.Forms.ProgressBar
    $PB.Name = "PowerShellProgressBar"
    $PB.Value = 0
    $PB.Style="Continuous"

    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 500 - 40
    $System_Drawing_Size.Height = 20
    $PB.Size = $System_Drawing_Size
    $PB.Left = 5
    $PB.Top = 40
    $ObjForm.Controls.Add($PB)
##############################################################
    $i = 0
    $timer = New-Object System.Windows.Forms.Timer 
    $timer.Interval = 1000

    $timer.add_Tick({
        StartProgressBar1
    })
    $timer.Enabled = $true
    $timer.Start()
##############################################################
    ## -- Show the Progress-Bar and Start The PowerShell Script
    $ObjForm.Show() | Out-Null
    $ObjForm.Focus() | Out-NUll
    $ObjLabel.Text = "Please wait ... While we create your Excel Spreadsheet "
    $ObjForm.Refresh()

    Start-Sleep -Seconds 1
    Out-Null
}

Function Folder {
$Form3.Dispose()
Import-Module $Ns1 -force
Get-permissions
}
Function Members {
$Form3.Dispose()
Import-Module $Ns2 -force
Get-Members
}
Function Shares{
$Form3.Dispose()
Import-Module $Ns3 -force
Get-Shares
}
Function Auto{
$Form3.Dispose()
Import-Module $Ns5 -force
Get-Auto
}
Function Get-Choice{
    Try{
   
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
    
    $Script:Form3 = New-Object System.Windows.Forms.Form
    
    $Form3.Text = "Please Make A Choice From Below"
    
    $Form3.ClientSize = New-Object System.Drawing.Size(390, 160)
    
    $form3.topmost = $true
    
    $Text = New-Object System.Windows.Forms.Label
    
    $Text.Location = New-Object System.Drawing.Point(15, 15)
    
    $Text.Size = New-Object System.Drawing.Size(300, 80)
    
    $Text.Text = "What would you like to do today? `n`n 1. Find the Permissions of a folder `n 2. Find Who Has Access to a Folder `n 3. Find out what shares are on a server`n 4. Select a csv that contains a list of shares"
    
    $Form3.Controls.Add($Text)
    
    #$ErrorActionPreference = "SilentlyContinue"
    
    Function Button1
    
    {
    
    $Button1 = New-Object System.Windows.Forms.Button
    
    $Button1.Location = New-Object System.Drawing.Point(15, 100)
    
    $Button1.Size = New-Object System.Drawing.Size(125, 25)
    
    $Button1.Text = "1. Folder Permissions"
    
    $Button1.add_Click({Folder
    
    $Form3.Close()})
    
    $Form3.Controls.Add($Button1)
    
    }
    
    Function Button2
    
    {
    
    $Button2 = New-Object System.Windows.Forms.Button
    
    $Button2.Location = New-Object System.Drawing.Point(145, 100)
    
    $Button2.Size = New-Object System.Drawing.Size(100, 25)
    
    $Button2.Text = "2. Folder Access"
    
    $Button2.add_Click({Members 
    
    $Form3.Close()})
    
    $Form3.Controls.Add($Button2)
    
    }
    
    Function Button3
    
    {
    
    $Button3 = New-Object System.Windows.Forms.Button
    
    $Button3.Location = New-Object System.Drawing.Point(250, 100)
    
    $Button3.Size = New-Object System.Drawing.Size(130, 25)
    
    $Button3.Text = "3. Shares on a Server"
    
    $Button3.add_Click({Shares 
    
    $Form3.Close()})
    
    $Form3.Controls.Add($Button3)
    
    }
    
    Function Button4
    
    {
    
    $Button4 = New-Object System.Windows.Forms.Button
    
    $Button4.Location = New-Object System.Drawing.Point(145, 128)
    
    $Button4.Size = New-Object System.Drawing.Size(100, 25)
    
    $Button4.Text = "4. Auto Shares"
    
    $Button4.add_Click({Auto})
    
    #$Form3.Close()}
    
    $Form3.Controls.Add($Button4)
    
    }
    
    Button1
    
    Button2
    
    Button3
    
    Button4

    [void]$Form3.showdialog()
}
Catch [System.Management.Automation.RuntimeException]{
    ER
$MSError =$_
$MSError2 = $_.exception.Message
$MSError| Out-File $ERL5 -append
$MSError2 | Out-File $ECL5 -append
Write-Host "Script Ended Prematurely in Main Script please see $ERL5" -ForegroundColor Magenta
Return
}
    }
    If($SZ -le "0"){
        Get-Scripts
    }
Get-Choice

}
Catch [System.Management.Automation.RuntimeException]{
ER
$MSError =$_
$MSError2 = $_.exception.Message
$MSError| Out-File $ERL5 -append
$MSError2 | Out-File $ECL5 -append
$Res = "Script Ended Prematurely in Main Script please see $ERL5"
$a = new-object -comobject wscript.shell
$b = $a.popup("$Res ",1," Error ",48)
#Write-Host "Script Ended Prematurely in Main Script please see $ERL5" -ForegroundColor Magenta
Return
#Write-Error $_
}