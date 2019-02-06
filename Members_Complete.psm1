#Write-Host "Members of Shares Loaded and ran"
$Script:ERL1 = $E4+ "Members_Errors_$Date.txt"
$Script:ECL1 = $E4 + "Members_Exceptions_$Date.txt"
Function Get-Members{
    $SZ = 1
Try{
##########################################################
########            Variables                     ######## 
##########################################################
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")
$Script:Ma = "C:\Temp\Server_Shares\Results\"
$Script:MB = "C:\Temp\Server_Shares\"
$Script:Vl = "2"
Function Check{
if( -Not (Test-Path -Path $Ma ) )

{
    $Vl = "0"
    $Res = "Please run the Folder Permission first"
    $a = new-object -comobject wscript.shell
    $b = $a.popup("$Res ",3," Error ",48)
    #Write-Host "Please run the Folder Permission first"
    Return
    }
    Else{
    $Vl = "1"
    Path
    }

    
}
Function Path{
    $openFileDialog = New-Object windows.forms.openfiledialog   
    $openFileDialog.initialDirectory = $Ma   
    $openFileDialog.title = "Please select the created Csv to Import"   
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
    #Write-Host "Cancelled by the user" -ForegroundColor Red
    $Res = "Cancelled by the user"
    $a = new-object -comobject wscript.shell
    $b = $a.popup("$Res ",3," Error ",48) 
    Return
    }
    Run
    Excel-Write

}


##########################################################
########            Main Function                 ######## 
##########################################################
Function Run{
Try{
$FB = $FS -replace ".{4}$"
#$FC = "User access for " + $FB
$Script:FP = $FS -replace "User access for ", ""
$FP = $FP -creplace '(?s)^.*\\', ''
#$FP = $FP -replace "User access for ", ""
$FP = $FP -replace ".{4}$"
$FP = $FP.trim()
$Script:res = $FP
$script:MB = $MB + $Res 
$Script:Ma2 = $MA + $Res
$Script:Ma3 = $MA + $Res +"\"
$Script:Ma4 = $MA + $Res +"\" + $Res
$Script:Ma5 = $MA + $Res +"\"
New-item -ItemType Directory -Path $Ma2
$Script:CSV =  $Ma4
$Script:CSV2 = $MA4 +"2"
$Script:UGS = $MA5 +"Members of group "
$Script:UNA = $MA5 +"Unable to Access"
$Script:UNA2 = $UNA + ".txt"
$Script:UNA3 = $UNA + ".csv"
Copy-Item "$FB.csv" "$CSV.csv"
Import-CSV "$CSV.csv" | Select -ExpandProperty "User Groups" | Out-file "$CSV.txt" # Imports csv that previous script created and changes it to a txt file

$Script:UG = GC "$CSV.txt"


$UG -creplace '(?s)^.*\\',"" | Set-Content "$CSV2.txt"  # Removes text that is not needed and saves the txt file 

$Script:UG2 = GC "$CSV2.txt"
Remove-item "$CSV.csv"
# Loops through the txt file and queries active directory
$UG2|ForEach-Object{

Try{
$Script:Group = $_
$FN = "$_"
$FN2 = $FN + ".txt"
$NC = "$Fn.csv"
$NC2 = $UGS + $NC
# Active Directory Query
$Query = Get-ADGroupMember -Identity $_ |`
Select SamAccountName, Name,`
@{Name="Title";Expression={(Get-ADUser $_.distinguishedName -Properties Title).title}},@{Name="Description";Expression={(Get-ADUser $_.distinguishedName -Properties Description).description}}  #| ft Name,SamAccountName,Title,Description -AutoSize 
$Query | Select-Object @{label = "Name";Expression = {$_.Name}},
@{label = "Account Name";Expression = {$_.SamAccountName}},
@{label = "Title";Expression = {$_.Title}},
@{label = "Description";Expression = {$_.Description}} | Export-csv $NC2 -NoClobber -NoTypeInformation
$N4 = $UGS + $FN2
$Co =$UG2.Length
If($Co -gt "1" ){
    While ($i -le $Hu) {
        $Hi =  $Hu / $Co
        $PB.Value = $i
        Start-Sleep -m 150
        $i
        $i += $Hi
        }
}
# If the script fails
}
Catch{
$God = "Unable to find users in Group $Group"

Write-host $God 
$God | Out-File $Una2 -Append
}
}
}
Catch{
    ER
    $MCError =$_
    $MCError2 = $_.exception.Message
    $MCError | Out-File $ERL1 -append
    $MCError2 | Out-File $ECL1 -append
    $Res = "Script Ended Prematurely in Members please see $ERL1"
    $a = new-object -comobject wscript.shell
    $b = $a.popup("$Res ",3," Error ",48)
}
}

##########################################################
########            Excel Function               ######## 
##########################################################
Function Excel-Write{
    If(Test-Path -Path $UNA2){
        Rename-item $UNA2 $UNA3
    }
    ProgressBar2
    $RD = $Ma3 + "*.csv" 
    $CsvDir = $RD 
    
    $csvs = dir -path $CsvDir # Collects all the .csv's from the driectory 
    $outputxls = "$Ma4.Xlsx"
    $Script:Excel = New-Object -ComObject excel.application
    $Excel.displayAlerts = $false
    $workbook = $excel.Workbooks.add()
    # Loops through each CVS, pulling all the data from each one
    foreach($iCsv in $csvs){
        $iCsv
        $WN = ($iCsv -Split "\\")[5]
        #$wn = ($WN -Split " ")[3]
        $WN = $WN -replace ".{5}$"
        $WN = $WN -replace "Members of group ",""
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
    Delete
}
##########################################################
########            Delete Function               ######## 
##########################################################
Function Delete{
    get-childitem $MA3 -recurse -force -include *.txt | remove-item -force #Removes all txt files from final directory
    get-childitem $MA3 -recurse -force -include *.csv | remove-item -force #Removes all CSV files from final directory
    $Res = "Your file is saved to"
    $Res2 = "$outputxls"
    $a = new-object -comobject wscript.shell
    $b = $a.popup("$Res `n`n$Res2 ",5," Information ",65)
    #Write-Host "Your file is saved to"-ForegroundColor White "`n`n$outputxls" -ForegroundColor Green
}
If ($Vl -eq "1"){
Run #Runs Main code
Excel-Write # Creates a readable excel file from the data
#Write-Host "Finished"
Import-Module $Ns4 -force
}
Else{
Path
}
}
Catch {
    If($_ -match "already exists"){
        $Res = "The file and directory currently exist" 
        $Res2 = "please remove and run the script again"
    $a = new-object -comobject wscript.shell
    $b = $a.popup("$Res `n`n$Res2 ",3," Error ",48)
        #Write-Host "The file and directory currently exist, please remove and run the script again" -ForegroundColor Yellow
    }
    $Excel.quit()
    ER
    $MCError =$_
    $MCError2 = $_.exception.Message
    $MCError | Out-File $ERL1 -append
    $MCError2 | Out-File $ECL1 -append
    $Res = "Script Ended Prematurely in Members please see $ERL1"
    $a = new-object -comobject wscript.shell
    $b = $a.popup("$Res ",3," Error ",48)
#Write-Host "Script Ended Prematurely in Members please see $ERL1" -ForegroundColor Magenta
Return
}
}

