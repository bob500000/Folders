#Write-Host "Folder Permissions Loaded and ran"
$Script:ERL6 = $E4+ "AutoFolder_Errors_$Date.txt"
$Script:ECL6 = $E4 + "AutoFolder_$Date.txt"
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")
$Script:Ma = "C:\Temp\Server_Shares\Server Results\"

Function Get-Auto{
Try{
    $SZ = 1
[void][System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")

[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")


$Script:Ma3 = "C:\Temp\Server_Shares\Server Results\"
    
Function Check{
    if( -Not (Test-Path -Path $Ma ) )
    
    {

        New-Item -ItemType directory -Path $MA |out-null
    
    }
    
        
    }
Function Path{
    Write-Host "Please minimize the console and select the csv you wish to use." - Foreground Yellow 
        $openFileDialog = New-Object windows.forms.openfiledialog   
        $openFileDialog.initialDirectory = $Ma   
        $openFileDialog.title = "Please select the Server you wish to find out about"   
        $openFileDialog.filter = "csv(*.csv)|*.csv"
        $openFileDialog.ShowHelp = $True   
        
        $ult = $result = $openFileDialog.ShowDialog((New-Object System.Windows.Forms.Form -Property @{TopMost = $true }))
    
        $ult 
    
        if($result -eq "OK")    {    
                #Write-Host "Selected Downloaded Settings File:"  -ForegroundColor Green  
                $Global:FS = $OpenFileDialog.filename 
                $Res = "File $FS Selected"
                $a = new-object -comobject wscript.shell
                $b = $a.popup("$Res ",3," Information ",65)
                 #Write-Host "File $FS Selected" -ForegroundColor DarkGreen
                }
        Else{
            $Res = "Cancelled by the user"
            $a = new-object -comobject wscript.shell
            $b = $a.popup("$Res ",3," Error ",48)    
        #Write-Host "Cancelled by the user" -ForegroundColor Red
        Return
        }
    }
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
            $Form12.Dispose()
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

        Return

     }

    }

  Function Get-Search{


Path

Import-Csv $FS


ProgressBar1
$FS| ForEach-Object{
$Res = (get-acl $_).Access  

$Script:Gold = $Res| Select-object @{label = "User Groups";Expression = {$_.IdentityReference}},
                            @{label = "Rights";Expression = {$_.FileSystemRights}},
                            @{label = "Access";Expression = {$_.AccessControlType}} 

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

Get-Question
$ObjForm.Close()
}
}

  Function Script:Get-Go{
    $Form1.Close()
$FPath = Select-FolderDialog

$Folder = $FPath + "\" + [Microsoft.VisualBasic.Interaction]::InputBox("Please select a folder to save the data to", "Path Choice") + "\"

[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null

"Please minimize the console to select a folder in which to save the results"

$Name = [Microsoft.VisualBasic.Interaction]::InputBox("Please choose a filename", "File Name Choice")

[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null

$cfgOutpath = $Folder + "$Name"

if( -Not (Test-Path -Path $Folder ) )

{

    New-Item -ItemType directory -Path $Folder |out-null

}

Else{

    [System.Windows.MessageBox]::Show('The directory already exists','Error','Ok','Error')

}

$Gold | Export-Csv "$cfgOutpath.csv" -NoClobber -NoTypeInformation

$Script:Ma3 = $cfgOutpath
$Res = "File has been saved to $cfgOutpath.csv"
$a = new-object -comobject wscript.shell
$b = $a.popup("$Res ",3," Information ",65)
#Write-Host "File has been saved to $cfgOutpath.csv" -ForegroundColor Yellow
Get-Q2
}
##############################################
##          Testing Phases                  ##
##            Get-Start                     ##
##            Search                        ##
##############################################


Function Script:Get-Create {
    $Form1.Close()
    if( -Not (Test-Path -Path $Ma3 ) )

{

    New-Item -ItemType directory -Path $Ma3 |out-null

}
Get-Done
}

Function Script:Get-Done{
Try{
$PC2 = ($FS -split '\\')[-1]
$PC2 = $PC2 -replace ".{5}$"

$CSV = "C:\Temp\Server_Shares\User access for $PC2"

$cfgOutpath = $CSV

if( -Not (Test-Path -Path "$cfgOutpath.csv" ) ){
    $Gold | Export-Csv "$cfgOutpath.csv" -NoClobber -NoTypeInformation

    $ObjForm.Close()
    Get-Q2
}
Else{

    $Res = "The File Already Exists" 
    $Res2 = "Please Delete and try again"
    $a = new-object -comobject wscript.shell
    $b = $a.popup("$Res `n`n$Res2 ",3," Error ",48)
    Return

}


}
Catch{
ER
$AFError =$_
$APError2 = $_.exception.Message
$AFError | Out-File $ERL6 -append
$AFError2 | Out-File $ECL6 -append
$Res = "Script Ended Prematurely in Folder Permissions please see $ECL6"
$a = new-object -comobject wscript.shell
$b = $a.popup("$Res ",3," Error ",48)
Return
}
}

Function Script:Get-Q2 {

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

$Button2.add_Click({Get-Excel

$Form2.Close()})

$Form2.Controls.Add($Button2)

}

Button1

Button2

[void]$form2.showdialog()
}


Function Script:Get-Ans{
    $Form2.Close()
    Try{
        Get-Excel
    }
    Catch{
    #Write-Host "Unable to create XSLX please check full path." -ForegroundColor Red
    $Res = "Unable to create XSLX please check full path."
    $a = new-object -comobject wscript.shell
    $b = $a.popup("$Res ",3," Error ",48)
    Return
    }
}

Function Script:Get-Result{
    $Form2.Dispose()
    $Res = "File has been saved to $CSV.csv"
    $a = new-object -comobject wscript.shell
    $b = $a.popup("$Res ",5," Information ",65)
}

Function Script:Get-Excel{
    ProgressBar2
    $RD = $Ma3 + "*.csv" 
    $CsvDir = $RD 
    $Ma4 = $cfgOutpath
    $csvs = dir -path $CsvDir # Collects all the .csv's from the driectory 
    $outputxls = "$Ma4.Xlsx"
    $Global:Excel = New-Object -ComObject excel.application
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
    #Write-Host "File has been saved to $outputxls" -ForegroundColor Yellow
    $ObjForm.Close()
    $Res = "Your file is saved to"
    $Res2 = "$outputxls"
    $a = new-object -comobject wscript.shell
    $b = $a.popup("$Res `n`n$Res2 ",5," Information ",65)
    Delete
}

Function Script:Delete{
    get-childitem $MA3 -recurse -force -include *.txt | remove-item -force #Removes all txt files from final directory
    get-childitem $MA3 -recurse -force -include *.csv | remove-item -force #Removes all CSV files from final directory
}

Get-Start

Write-Host "Finished"
Import-Module $Ns4 -force
}
Catch {
ER
$Excel.quit()
$AFError =$_
$APError2 = $_.exception.Message
$AFError | Out-File $ERL6 -append
$AFError2 | Out-File $ECL6 -append
$Res = "Script Ended Prematurely in Auto Folders please see $ERL6"
$a = new-object -comobject wscript.shell
$b = $a.popup("$Res ",3,3," Error ",48)
#Write-Host "Script Ended Prematurely in Auto Folders please see $ERL2" -ForegroundColor Magenta
Return
}
}







