#Write-Host "Folder Permissions Loaded and ran"
$Script:ERL2 = $E4+ "Folder_Errors_$Date.txt"
$Script:ECL2 = $E4 + "Folder_Exceptions_$Date.txt"
$Script:Val = 0
Function Get-Permissions{
Try{
[void][System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")

[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")


$Script:Ma3 = "C:\Temp\Server_Shares\"
Try{
Function Get-End{
$Form1.Dispose()
$Form12.Dispose()
   Return
}
Function Get-Stop{
    $Form1.Dispose()
    Restart
    }    
Function Get-Sc{
    $Global:SZ = 1
    Import-Module $E3 -force
    $Form12.Dispose()
}
Function Restart {
$Global:SA = 1
        [void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        [void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
        $Script:Form12 = New-Object System.Windows.Forms.Form
        
        $Form12.ClientSize = New-Object System.Drawing.Size(220, 100)
        
        $form12.topmost = $true
        
        $Text = New-Object System.Windows.Forms.Label
        
        $Text.Location = New-Object System.Drawing.Point(15, 15)
        
        $Text.Size = New-Object System.Drawing.Size(200, 40)
        
        $Text.Text = "Would you like to run another function?"
        
        $Form12.Controls.Add($Text)
        
        Function Button1
        
        {
        
        $Button1 = New-Object System.Windows.Forms.Button
        
        $Button1.Location = New-Object System.Drawing.Point(20, 55)
        
        $Button1.Size = New-Object System.Drawing.Size(55, 20)
        
        $Button1.Text = "Yes"
        
        $Button1.add_Click({Get-SC -ErrorAction SilentlyContinue
        
        $Form12.Close()})
        
        $Form12.Controls.Add($Button1)
        
        }
        
        Function Button2
        
        {
        
        $Button2 = New-Object System.Windows.Forms.Button
        
        $Button2.Location = New-Object System.Drawing.Point(80, 55)
        
        $Button2.Size = New-Object System.Drawing.Size(55, 20)
        
        $Button2.Text = "No"
        
        $Button2.add_Click({Get-End
        
        $Form12.Close()})
        
        $Form12.Controls.Add($Button2)
        }
        
        Button1
        
        Button2
        
        [void]$form12.showdialog()
        
        }
Function Rerun {
$Script:Ez = 0 
        [void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        [void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
        $Script:E = 1
        $Script:Form1 = New-Object System.Windows.Forms.Form
        
        $Form1.ClientSize = New-Object System.Drawing.Size(200, 100)
        
        $form1.topmost = $true
        
        $Text = New-Object System.Windows.Forms.Label
        
        $Text.Location = New-Object System.Drawing.Point(15, 15)
        
        $Text.Size = New-Object System.Drawing.Size(200, 40)
        
        $Text.Text = "Do you want to try another set of credentials?"
        
        $Form1.Controls.Add($Text)
        
        Function Button1
        
        {
        
        $Button1 = New-Object System.Windows.Forms.Button
        
        $Button1.Location = New-Object System.Drawing.Point(20, 55)
        
        $Button1.Size = New-Object System.Drawing.Size(55, 20)
        
        $Button1.Text = "Yes"
        
        $Button1.add_Click({Get-Start -ErrorAction SilentlyContinue
        
        $Form1.Close()})
        
        $Form1.Controls.Add($Button1)
        
        }
        
        Function Button2
        
        {
        
        $Button2 = New-Object System.Windows.Forms.Button
        
        $Button2.Location = New-Object System.Drawing.Point(80, 55)
        
        $Button2.Size = New-Object System.Drawing.Size(55, 20)
        
        $Button2.Text = "No"
        
        $Button2.add_Click({Get-Stop
        
        $Form1.Close()})
        
        $Form1.Controls.Add($Button2)
        }
        
        Button1
        
        Button2
        
        [void]$form1.showdialog()
        
        }
Function Checking {
            $Info = "Please Wait whilst we confirm your access"
            $Info2 = "Checking.."
            $a = new-object -comobject wscript.shell
            $b = $a.popup("$Info `n$Info2 ",1," Information ",65)
            $Info = "Please Wait whilst we confirm your access"
            $Info2 = "Checking...."
            $a = new-object -comobject wscript.shell
            $b = $a.popup("$Info `n$Info2 ",1," Information ",65)
            $Info = "Please Wait whilst we confirm your access"
            $Info2 = "Checking......"
            $a = new-object -comobject wscript.shell
            $b = $a.popup("$Info `n$Info2 ",1," Information ",65)
            $Info = "Please Wait whilst we confirm your access"
            $Info2 = "Checking........"
            $a = new-object -comobject wscript.shell
            $b = $a.popup("$Info `n$Info2 ",1," Information ",65)
            $Info = "Please Wait whilst we confirm your access"
            $Info2 = "Checking.........."
            $a = new-object -comobject wscript.shell
            $b = $a.popup("$Info `n$Info2 ",1," Information ",65)
            }
Function Get-Start #Prompts User for Domain Admin Account details
{
    Try{
        If($E -ge "1")
        {
            $Form1.Dispose()
        }
        $Val ++
###################################################################################
###################################################################################
#
#       Requests User for admin details
#
###################################################################################
###################################################################################    
    #Get user credentials 
    $Cred = Get-Credential  -Message "Enter Your Domain Admin Credentials (Domain\Username)" -UserName Swinton\
    if ($Cred -eq $Null) 
                        { 
                            $Res1 = "Please enter your username in the form of Swinton\UserName and try again"
                            Write-Host $Res1 -BackgroundColor Black -ForegroundColor Yellow  
                            $a = new-object -comobject wscript.shell
                            $b = $a.popup("$Res1 ",3," Error ",16)
                            Rerun
                            Return                          
                        } 
                        Checking
    #Parse provided user credentials 
    $DomainNetBIOS = $Cred.username.Split("{\}")[0] 
    $UserName = $Cred.username.Split("{\}")[1] 
    $Password = $Cred.GetNetworkCredential().password 
     
    Write-Host "`n" 
    Write-Host "Checking Credentials for $DomainNetBIOS\$UserName" -BackgroundColor Black -ForegroundColor White 
    Write-Host "***************************************" 
 
    If ($DomainNetBIOS -eq $Null -or $UserName -eq $Null)  
                        { 
                            $Res2 = "Missing domain please type in the following format: Domain\Username"
                            #Write-Host $Res2 -BackgroundColor Black -ForegroundColor Yellow
                            $a = new-object -comobject wscript.shell
                            $b = $a.popup("$Res2 ",5," Error ",16) 
                            Rerun 
                            Return
                        } 
    #    Checks if the domain in question is reachable, and get the domain FQDN. 
    Try 
    { 
        $DomainFQDN = (Get-ADDomain $DomainNetBIOS).DNSRoot 
    } 
    Catch 
    { 
        $Res3 = "Error: Domain was not found: "
        $Res33=$_.Exception.Message
        $Res333 = "Please make sure the domain NetBios name is correct, and is reachable from this computer"
        #Write-Host $Res3  -BackgroundColor Black -ForegroundColor Red 
        #Write-Host $Res33 -BackgroundColor Black -ForegroundColor Red 
        $a = new-object -comobject wscript.shell
        $b = $a.popup("$Res3 `n$Res33`n$Res333 ",5," Error ",16)
        Rerun 
        Return
    } 
     
    #Checks user credentials against the domain 
    $DomainObj = "LDAP://" + $DomainFQDN 
    $DomainBind = New-Object System.DirectoryServices.DirectoryEntry($DomainObj,$UserName,$Password) 
    $DomainName = $DomainBind.distinguishedName 
     
    If ($DomainName -eq $Null) 
        { 
            Write-Host "Domain $DomainFQDN was found: True" -BackgroundColor Black -ForegroundColor Green 
         
            $UserExist = Get-ADUser -Server $DomainFQDN -Properties LockedOut -Filter {sAMAccountName -eq $UserName} 
            If ($UserExist -eq $Null)  
                        { 
                            $Res4 = "Error: Username $Username does not exist in $DomainFQDN Domain."
                            Write-Host $Res4 -BackgroundColor Black -ForegroundColor Red 
                            $a = new-object -comobject wscript.shell
                            $b = $a.popup("$Res4 ",5," Error ",16)
                            Rerun 
                            Return 
                        } 
            Else  
                        {    
                            Write-Host "User exists in the domain: True" -BackgroundColor Black -ForegroundColor Green 
 
 
                            If ($UserExist.Enabled -eq "True") 
                                    { 
                                        Write-Host "User Enabled: "$UserExist.Enabled -BackgroundColor Black -ForegroundColor Green 
                                    } 
 
                            Else 
                                    { 
                                        $Res5 = "User Enabled: " + $UserExist.Enabled
                                        $Res55 = "Enable the user account in Active Directory, Then check again"
                                        Write-Host $Res5 -BackgroundColor Black -ForegroundColor RED 
                                        Write-Host $Res55 -BackgroundColor Black -ForegroundColor RED 
                                        $a = new-object -comobject wscript.shell
                                        $b = $a.popup("$Res5 `n$Res55 ",5," Error ",16)
                                        Rerun 
                                        Return 
                                    } 
 
                            If ($UserExist.LockedOut -eq "True") 
                                    { 
                                        $Res6 = "User Locked: "+ $UserExist.LockedOut
                                        $Res65 = "Unlock the User Account in Active Directory, Then check again..."
                                        Write-Host $Res6 -BackgroundColor Black -ForegroundColor Red 
                                        Write-Host $Res66 -BackgroundColor Black -ForegroundColor RED 
                                        $a = new-object -comobject wscript.shell
                                        $b = $a.popup("$Res6 `n$Res65 ",5," Error ",16)
                                        Rerun 
                                        Return
                                    } 
                            Else 
                                    { 
                                        $Res8 = "Authentication failed for"
                                        $Res88 = "$DomainNetBIOS\$UserName with the provided password."
                                        $Res888 = "Please confirm the password, and try again..."
                                        #Write-Host $Res8 -BackgroundColor Black -ForegroundColor Red 
                                        #Write-Host $Res88 -BackgroundColor Black -ForegroundColor Red 
                                        $a = new-object -comobject wscript.shell
                                        $b = $a.popup("$Res8 `n$Res88`n$Res888 ",5," Error ",16)
                                        Rerun 
                                        Return
                                    } 
                        } 
            
            $Res8 = "Authentication failed for $DomainNetBIOS\$UserName with the provided password."
            $Res88 = "Please confirm the password, and try again..."
            #Write-Host $Res8 -BackgroundColor Black -ForegroundColor Red 
            #Write-Host $Res88 -BackgroundColor Black -ForegroundColor Red 
            $a = new-object -comobject wscript.shell
            $b = $a.popup("$Res8 `n$Res88 ",5," Error ",16)
            Rerun 
            Return
        } 
      
    Else 
        { 
            $Res9 = "SUCCESS: The account $Username" 
            $Res99 = "Successfully authenticated against the domain: $DomainFQDN"
            Write-Host $Res9 -BackgroundColor Black -ForegroundColor Green 
            $a = new-object -comobject wscript.shell
            $b = $a.popup("$Res9 `n`n$Res99",3," Success",65)
        Get-Search
        } 
    }
    Catch [System.Management.Automation.ActionPreferenceStopException]{
        ER
        $FPError =$_
        $FPError2 = $_.exception.Message
        $FPError | Out-File $ERL2 -append
        $FPError2 | Out-File $ECL2 -append
        $ErrorM = "Script Ended Prematurely in Folder Permissions please see $ERL2"
        $a = new-object -comobject wscript.shell
            $b = $a.popup("$ErrorM",5," Error ",1)
        Return

    }
}
###################################################################################
###################################################################################
}
Catch{
    ER
        $FPError =$_
        $FPError2 = $_.exception.Message
        $FPError | Out-File $ERL2 -append
        $FPError2 | Out-File $ECL2 -append
        $Res = "Script Ended Prematurely in Folder Permissions please see $ERL2"
        $a = new-object -comobject wscript.shell
        $b = $a.popup("$Res ",3," Information ",48)
        Return
}
  Function Script:Get-Question {
$Val ++
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

$Form1 = New-Object System.Windows.Forms.Form

$Form1.ClientSize = New-Object System.Drawing.Size(200, 100)

$form1.topmost = $true

$Text = New-Object System.Windows.Forms.Label

$Text.Location = New-Object System.Drawing.Point(15, 15)

$Text.Size = New-Object System.Drawing.Size(200, 40)

$Text.Text = "Would you like to save the file to a custom location?"

$Form1.Controls.Add($Text)

#$ErrorActionPreference = "SilentlyContinue"

Function Button1

{

$Button1 = New-Object System.Windows.Forms.Button

$Button1.Location = New-Object System.Drawing.Point(20, 55)

$Button1.Size = New-Object System.Drawing.Size(55, 20)

$Button1.Text = "Yes"

$Button1.add_Click({Get-Go -ErrorAction SilentlyContinue

$Form1.Close()})

$Form1.Controls.Add($Button1)

}

Function Button2

{

$Button2 = New-Object System.Windows.Forms.Button

$Button2.Location = New-Object System.Drawing.Point(80, 55)

$Button2.Size = New-Object System.Drawing.Size(55, 20)

$Button2.Text = "No"

$Button2.add_Click({Get-Create -ErrorAction SilentlyContinue

$Form1.Close()})

$Form1.Controls.Add($Button2)

}

Button1

Button2

[void]$form1.showdialog()

}

  Function Select-FolderDialog{
    
    

         param([string]$Description="Select Folder",[string]$RootFolder="Desktop")



     [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null     
     $Val ++
     Write-host "Please minimize the console to select a folder in which to save the results"

     $objForm = New-Object System.Windows.Forms.FolderBrowserDialog

     $objForm.Rootfolder = $RootFolder

     $objForm.Description = $Description

     $objForm.ShowNewFolderButton = $false

     $Show = $objForm.ShowDialog()

     If ($Show -eq "OK")

     {

         Return $objForm.SelectedPath

     }

     Else

     {
        $Res = "Operation cancelled by user."
        $a = new-object -comobject wscript.shell
        $b = $a.popup("$Res ",3," Error ",48)
        #Write-Error "Operation cancelled by user."

        Return

     }

    }

  Function Get-Search{
Try{
    $Val ++
Write-host $Res -ForegroundColor Green
    $Res = "Please Minimize the Windows and enter the full folder path that you require permissions for"
    $a = new-object -comobject wscript.shell
    $b = $a.popup("$Res ",3," Information ",65)

$Script:PC = [Microsoft.VisualBasic.Interaction]::InputBox("Please enter the full path of the folder you wish to search", "Folder choice")


If ($PC -eq "")

{
    $Res = "No Path Entered."
    $Res2 = "Ending Application."
    $a = new-object -comobject wscript.shell
    $b = $a.popup("$Res`n`n$Res2 ",3," Error ",48)
Return

}

ProgressBar1
Try{
    $Res = dir $PC | Get-Acl | Select Path,Access -Exp Access

    $Script:Gold = $Res | Select-Object @{Label="Path";Expression={Convert-Path $_.Path}},
                                        @{Label="User";Expression={$_.IdentityReference}},
                                        @{Label="Access";Expression={$_.FileSystemRights}}
<#    
$Res = (get-acl $pc).Access 

$Script:Gold = $Res| Select-object @{label = "User Groups";Expression = {$_.IdentityReference}},
                            @{label = "Rights";Expression = {$_.FileSystemRights}},
                            @{label = "Access";Expression = {$_.AccessControlType}} #>
$Co =$Res.Length
If($Co -gt "1" ){
    While ($i -le $Hu) {
        $Hi =  $Hu / $Co
        $PB.Value = $i
        Start-Sleep -m 150
        $i
        $i += $Hi
        }
}
$ObjForm.Close()
Get-Question
}
Catch {
    If($_ -match "Attempted to perform an unauthorized operation."){
    ER
    $ObjForm.Close()
    $ErrorU = $_
    $ErrorsU = $_.exception.Message
    $ErrorU | Out-File $ERL2 -append
    $ErrorsU | Out-File $ECL2 -append
    $Res2 = "Access Denied"
    $Res = "Please confirm you have the relevant Access"
    $a = new-object -comobject wscript.shell
    $b = $a.popup("$Res `n`n$Res2",3," Information ",65)
    Return
    }
    ER
    $ObjForm.Close()
    $Errors = $_
    $Errors2 = $_.exception.Message
    $Errors | Out-File $ERL2 -append
    $Errors2 | Out-File $ECL2 -append
    $Res2 = "Path not found"
    $Res = "Please confirm you have entered the correct path"
    $a = new-object -comobject wscript.shell
    $b = $a.popup("$Res `n`n$Res2",3," Information ",65)
    Return
}
}
Catch {
$ObjForm.Close()
ER
$FPError =$_
$FPError2 = $_.exception.Message
$FPError | Out-File $ERL2 -append
$FPError2 | Out-File $ECL2 -append
$Res = "Script Ended Prematurely in Folder Permissions please see $ERL2"
$a = new-object -comobject wscript.shell
$b = $a.popup("$Res ",3," Information ",65)
Return
}
}

  Function Script:Get-Go{
    $Val ++
    $Form1.Dispose()
$FPath = Select-FolderDialog

$Res = "Please minimize the console to select a folder in which to save the results"
$a = new-object -comobject wscript.shell
$b = $a.popup("$Res ",2," Information ",65)

$Folder = $FPath + "\" + [Microsoft.VisualBasic.Interaction]::InputBox("Please Choose a folder to save the data to", "Path Choice") + "\"

[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null

$Res = "Please minimize the console " 
$Res2 = "Choose a name for the  folder in which to save the results"
$a = new-object -comobject wscript.shell
$b = $a.popup("$Res `n$Res2 ",2," Information ",65)

$Name = [Microsoft.VisualBasic.Interaction]::InputBox("Please choose a filename", "File Name Choice")

[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null

$CFG = $Folder + "Delete_ME"
$cfgOutpath = $Folder + "$Name"

if( -Not (Test-Path -Path $Folder ) )

{

    New-Item -ItemType directory -Path $Folder |out-null

}

Else{

    [System.Windows.MessageBox]::Show('The directory already exists','Error','Ok','Error')

}

$Gold | Export-Csv "$CFG.csv" -NoClobber -NoTypeInformation

$Import = Import-csv "$CFG.csv"
$Import | Foreach-Object{
$T = ""
$_.path = $_.path -replace $PC2,$T
} 


$Import | Export-Csv "$cfgOutpath.csv" -NoClobber -NoTypeInformation

$Script:Ma3 = $cfgOutpath

$Res = "File has been saved to $cfgOutpath.csv"
$a = new-object -comobject wscript.shell
$b = $a.popup("$Res ",3," Information ",65)
Get-Q2
}

##############################################
##          Testing Phases                  ##
##            Get-Start                     ##
##            Search                        ##
##############################################


Function Script:Get-Create {
    $Val ++
    $Form1.Dispose()
    if( -Not (Test-Path -Path $Ma3 ) )

{

    New-Item -ItemType directory -Path $Ma3 |out-null

}
Get-Done
}

Function Script:Get-Done{
Try{
    $Val ++
$PC2 = ($PC -split '\\')[-1]

$CSV = "C:\Temp\Server_Shares\User access for $PC2"
$CFG = "C:\Temp\Server_Shares\Delete_Me"

$cfgOutpath = $CSV

if( -Not (Test-Path -Path "$cfgOutpath.csv" ) ){
    
$Gold | Export-Csv "$CFG.csv" -NoClobber -NoTypeInformation

$Import = Import-csv "$CFG.csv"
$Import | Foreach-Object{
$T = ""
$_.path = $_.path -replace $PC2,$T
} 

$Import | Export-Csv "$cfgOutpath.csv" -NoClobber -NoTypeInformation

$ObjForm.Close()
Remove-Item "$CFG.csv"
Get-Q2
}
Else{
    $Res = "The File Already Exists" 
    $Res2 = "Please Delete and try again"
    $a = new-object -comobject wscript.shell
    $b = $a.popup("$Res `n`n$Res2 ",3," Error ",48)
    #[System.Windows.MessageBox]::Show('The File Already Exists Please Delete','Error','Ok','Error')
    Return
}

}
Catch {
ER
$FPError =$_
$FPError2 = $_.exception.Message
$FPError | Out-File $ERL2 -append
$FPError2 | Out-File $ECL2 -append
$Res = "Script Ended Prematurely in Folder Permissions please see $ERL2"
$a = new-object -comobject wscript.shell
$b = $a.popup("$Res ",3," Error ",48)
Return
}
}

Function Script:Get-Q2 {
    $Val ++
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

$Form2 = New-Object System.Windows.Forms.Form

$Form2.ClientSize = New-Object System.Drawing.Size(200, 100)

$form2.topmost = $true

$Text = New-Object System.Windows.Forms.Label

$Text.Location = New-Object System.Drawing.Point(15, 15)

$Text.Size = New-Object System.Drawing.Size(200, 40)

$Text.Text = "Would you like to create an Xlsx document or leave it as csv?"

$Form2.Controls.Add($Text)

$ErrorActionPreference = "SilentlyContinue"

Function Button1

{

$Button1 = New-Object System.Windows.Forms.Button

$Button1.Location = New-Object System.Drawing.Point(20, 55)

$Button1.Size = New-Object System.Drawing.Size(55, 20)

$Button1.Text = "CSV"

$Button1.add_Click({Get-Result

$Form2.Close()})

$Form2.Controls.Add($Button1)}


Function Button2

{

$Button2 = New-Object System.Windows.Forms.Button

$Button2.Location = New-Object System.Drawing.Point(80, 55)

$Button2.Size = New-Object System.Drawing.Size(55, 20)

$Button2.Text = "XLSX"

$Button2.add_Click({Get-Ans

$Form2.Close()})

$Form2.Controls.Add($Button2)

}

Button1

Button2

[void]$form2.showdialog()
}


Function Script:Get-Ans{
    $Val ++
    $Form2.Dispose()
    Try{
        Get-Excel
    }
    Catch{
        #Write-Host "Unable to create XSLX please check full path." -ForegroundColor Red
    $Excel.quit()
    ER
    $FPError =$_
    $FPError2 = $_.exception.Message
    $FPError | Out-File $ERL2 -append
    $FPError2 | Out-File $ECL2 -append
    $Res = "Unable to create XSLX please check full path."
    $a = new-object -comobject wscript.shell
    $b = $a.popup("$Res ",3," Error ",48)
        Return
    }
}

Function Script:Get-Result{
    $Val ++
    $Form2.Dispose()
    $Res = "File has been saved to $CSV.csv"
    $a = new-object -comobject wscript.shell
    $b = $a.popup("$Res ",5," Information ",65)
#Write-Host "File has been saved to $CSV.csv" -ForegroundColor Yellow
}

Function Script:Get-Excel{
    Try{
        $Val ++
    $Form2.Dispose()
    ProgressBar2
    $RD = $Ma3 + "*.csv" 
    $CsvDir = $RD 
    $Ma4 = $cfgOutpath
    $csvs = dir -path $CsvDir # Collects all the .csv's from the driectory 
    $outputxls = "$Ma4.Xlsx"
    $Script:Excel = New-Object -ComObject excel.application
    $Excel.displayAlerts = $false
    $workbook = $excel.Workbooks.add()
    # Loops through each CVS, pulling all the data from each one
    foreach($iCsv in $csvs){
        $iCsv
        $WN = ($iCsv -creplace '(?s)^.*\\', '')
        $WN = $WN -replace ".{4}$"
        If($WN.length -gt 30){
            $WN = $WN.Substring(0, [Math]::Min($WN.Length, 20))
            }
        $Worksheet = $workbook.worksheets.add()
        $Worksheet.name = $WN

        $TxtConnector = ("TEXT;" + $iCsv)
        $Connector = $worksheet.Querytables.add($txtconnector,$worksheet.Range("A1"))
        $query = $Worksheet.QueryTables.item($Connector.name)

        $query.TextfileOtherDelimiter = $Excel.Application.International(5)

        $Query.TextfileParseType =1
        $Query.TextFileColumnDataTypes = ,2 * $worksheet.cells.column.count
        $query.AdjustColumnWidth =1

        $Query.Refresh()
        $Query.Delete()
        $Worksheet.Cells.EntireColumn.AutoFit()
        $Worksheet.Rows.Item(1).Font.Bold = $true
        $Worksheet.Rows.Item(1).HorizontalAlignment = -4108
        $Worksheet.Rows.Item(1).Font.Underline = $true
        $Workbook.save()
        $Ca =$csvs.Length
        If($Ca -le "1"){
            Continue
        }
        Else{
        While ($i -le $Hu) {
        $Hi =  $Hu / $CA
        $PB.Value = $i
        Start-Sleep -m 150
        $i
        $i += $Hi
        }
                }
        }
    
    $Empty = $workbook.worksheets.item("Sheet1")
    $Empty.Delete()
    $Workbook.SaveAs($outputxls,51)
    $Workbook.close()
    $Excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    $ObjForm.Close()
    $Res = "Your file is saved to" 
    $Res2 = "$outputxls"
    $a = new-object -comobject wscript.shell
    $b = $a.popup("$Res `n`n$Res2 ",5," Information ",65)
    #Write-Host "Your file is saved to"-ForegroundColor White "`n`n$outputxls" -ForegroundColor Green
    Delete
    }
    Catch {
        ER
        $FPError =$_
        $FPError2 = $_.exception.Message
        $FPError | Out-File $ERL2 -append
        $FPError2 | Out-File $ECL2 -append
        $Res = "Script Ended Prematurely in Folder Permissions please see $ERL2"
        $a = new-object -comobject wscript.shell
        $b = $a.popup("$Res ",3," Information ",48)
        Return
    }
}

Function Script:Delete{
    $Val ++
    get-childitem $MA3 -recurse -force -include *.txt | remove-item -force #Removes all txt files from final directory
    get-childitem $MA3 -recurse -force -include *.csv | remove-item -force #Removes all CSV files from final directory
}

Get-Start

#Write-Host "Finished"
Import-Module $Ns4 -force
}
Catch {
    If($Val -gt "5"){
        $Excel.quit()
    }
ER
$FPError =$_
$FPError2 = $_.exception.Message
$FPError | Out-File $ERL2 -append
$FPError2 | Out-File $ECL2 -append
$Res = "Script Ended Prematurely in Folder Permissions please see $ERL2"
$a = new-object -comobject wscript.shell
$b = $a.popup("$Res ",3," Information ",48)
Return
}
}







