#Write-Host "Shares Loaded and ran"
$Script:erroractionpreference = 'Stop'
$Script:Var = "0"
$Script:ERL3 = $E4 + "Shares_Errors_$Date.txt"
$Script:ECL1 = $E4 + "Shares_Exceptions1_$Date.txt"
$Script:ERL1 = $E4 + "Shares_Errors1_$Date.txt"
$Script:ECL3 = $E4 + "Shares_Exceptions_$Date.txt"
$Global:V = 0
Function Get-Shares{
    $SZ = 1
Function Get-End{
    $Res = "Please Delete the existing folder and try again"
    $a = new-object -comobject wscript.shell
    $b = $a.popup("$Res ",3," Error ",48)
    #Write-Host "Please Delete the existing folder and try again"
    #$ObjForm.Close()
    Return
}
##############################################################

#Varibles

##############################################################
Try{

$Script:SP = "C:\Temp\Servers\"

$Script:TP = "C:\Temp\Servers\Pc.txt"
$Script:FSCSV = "C:\Temp\Server_Shares\Server Lists\"
$Script:Message1 = "Unknown Hosts"
$Script:Message2 = "Unable to connect"
$Script:Message3 = "Unknown Errors Occurred"
$Script:Txt = ".txt"
$Script:OT = ".csv"
$Script:FSERROR1 = $FSCSV+$Message1+$OT
$Script:FSERROR2 = $FSCSV+$Message2+$OT
$Script:FSERROR3 = $FSCSV+$Message2+$OT




#############################################################

#############################################################

Function Location{

    if( -Not (Test-Path -Path $SP ) ){

    New-Item -ItemType directory -Path $SP |out-null

}
}
Function File{

    if( -Not (Test-Path -Path $TP) ){
    $Res = "Please Create a text file with a list of servers and save it to; `n`n $TP"
    $Res2 = "Ending Application."
    $a = new-object -comobject wscript.shell
    $b = $a.popup("$Res`n`n$Res2 ",3," Error ",48)
    #Write-Host "`nPlease Create a text file with a list of servers and save it to; `n`n $TP" -ForegroundColor Yellow|out-null
    return

}
Shares

}

Function Shares{
Try{

    $PCs = gc "C:\Temp\Servers\Pc.txt"
ProgressBar1
if( Test-Path -Path $FSCSV ) 
{
    $Global:V ="1"
    Get-End
    
}
Else{
        New-Item -ItemType directory -Path $FSCSV |out-null

$PCs|Foreach-Object{ 
    If($V -eq "1")
    {
        Get-End
    }
Try{
    #$ErrorActionPreference ='Stop'
    $Script:PC = $_
    #Get-WmiObject win32_share -ComputerName $_ 
    #$DNSCheck = ([System.Net.Dns]::GetHostByName(($_)))

    $FSCSV2 = $FSCSV + $_
    $Script:Se = Get-WmiObject win32_share -ComputerName $_ | Sort-Object -Property path | out-null
    
    $se| Select-object @{Name="Server";Expression={$_.__Server}},
                       @{Name="Path";Expression={$_.path}},
                       @{Name="Name";Expression={$_.name}}  | Export-csv "$FSCSV2.csv" -NoClobber -NoTypeInformation |out-null 
    
    Get-WmiObject win32_share -ComputerName $_ 
    
    $DNSCheck = ([System.Net.Dns]::GetHostByName(($_))) 

    $Check = $DNSCheck.hostname

    $CH1 = $Check.split(".")[1]
    $CH2 = $Check.split(".")[2]
    $CH3 = $Check.split(".")[3]

    $Script:Ch4 =$CH1 +"." + $CH2 +"."+ $CH3

    #$CH4
    $Co =$PCs.Length
    If($Co -le "1"){
        Continue
    }
    Else{
    While ($i -le $Hu) {
    $Hi =  $Hu / $Co
    $PB.Value = $i
    Start-Sleep -m 150
    $i
    $i += $Hi
    }
    }
    }
  Catch{ 
    If($_ -match "No Such Host is known"){
        $PCx = $FSCSV + $PC + $OT
        If(Test-path  $PCx){
            Remove-item $PCx
        }
        ####################################
        #       Custom Csv File            #
        ####################################
        New-Object PSObject -Property @{
            Server = $Pc
            Error = $_.exception.InnerException.Message
        } | Export-CSV $FSERROR1 -append -NoTypeInformation -NoClobber
        $ErrorActionPreference ='Continue'
        ER
        $ER1 = $_
        $ER1 | Out-file $ERL1 -Append
    }  
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    Elseif ($_ -match "The RPC Server is unavailable"){
        New-Object PSObject -Property @{
            Server = $Pc
            Error = "Unable to connect"
        } | Export-CSV $FSERROR2 -append -NoTypeInformation -NoClobber
    ER
        $ER2 = $_
        $ER2 | Out-file $ERL1 -Append
   $ErrorActionPreference ='Continue'
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    else {
        New-Object PSObject -Property @{
        Server = $Pc
        Error = $_.exception.Message
        } | Export-CSV $FSERROR3 -append -NoTypeInformation -NoClobber
        $ER2 = $_
        $ER2 | Out-file $ERL1 -Append
   $ErrorActionPreference ='Continue'
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
}
}
$ObjForm.Close()
Excel-Write
#Return
}
Catch 
{
    IF($_ -Match "File does not exist"){
        $Res = "`nUnable to find C:\Temp\Servers\Pc.txt `n`n Please create the file and try again"
        $a = new-object -comobject wscript.shell
        $b = $a.popup("$Res ",3," Error ",48)
        #Write-Host "`nUnable to find C:\Temp\Servers\Pc.txt `n`n Please create the file and try again" -ForegroundColor Yellow 
    }
ER
$ObjForm.Close()
$SCError0 =$_
$SCError1 = $SCError0.exception.Message
$SCError0 | Out-File $ERL1 -append
$SCError1 | Out-File $ECL1 -append
#Write-Host "Script Ended Prematurely please see $ERL3" -ForegroundColor Magenta
$Res = "Script Ended Prematurely please see $ERL3"
$a = new-object -comobject wscript.shell
$b = $a.popup("$Res ",3," Error ",48)
#Write-Error $_
Return

}
}


#############################################################

#############################################################

Function Excel-Write{
    If($V -eq "1")
    {
        Return
    }

    #####################################################
    #         Multiple File Extension Re-name           #
    #
    #$TestP = Get-ChildItem  $FSCSV -recurse -filter $Txt
    #If($TestP -like $txt){
    #    Get-ChildItem -Path $FSCSV -Filter $Txt| Rename-Item -NewName {[System.IO.Path]::ChangeExtension($_.Name, $Ot)}
    #}
    ######################################################
    ProgressBar2
    $RD = $FSCSV + "*.csv" 
    $CsvDir = $RD 
    $Ma4 = $FSCSV + "All Server Shares for Domain $CH4"
    $csvs = dir -path $CsvDir # Collects all the .csv's from the driectory 
    $FSh = $csvs | Select-Object -First 1
    $FSh = ($FSh -Split "\\")[4]
    $FSh = $FSh -replace ".{5}$"
    $FSh
    $outputxls = "$Ma4.xlsx"
    $Script:Excel = New-Object -ComObject excel.application
    $Excel.displayAlerts = $false
    $workbook = $excel.Workbooks.add()
    # Loops through each CVS, pulling all the data from each one
   $RD | Foreach-object{
    Foreach($iCsv in $csvs){
        $Script:iCsv
        $WN = ($iCsv -Split "\\")[-1]
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

Function Delete{
    If($V -eq "1")
    {
        Return
    }
    $erroractionpreference = 'Stop'
        $Cop = $FSCSV + "*.csv"
        $FSCVSD = $FSCSV + "Server_CSVs"
        New-Item -ItemType directory -Path $FSCVSD |out-null
        Copy-Item $Cop -recurse $FSCVSD -force
        Remove-Item $Cop -force
        #Write-Host "Your file is saved to"-ForegroundColor White 
        #Write-Host "`n`n$outputxls" -ForegroundColor Green
        $Res = "Your file is saved to"
        $Res2 = "`n`n$outputxls"
        $a = new-object -comobject wscript.shell
        $b = $a.popup("$Res`n`n$Res2 ",3," Information ",65)
    }

#############################################################

#############################################################
Location
File
Import-Module $Ns4 -force
}

Catch 
{
    $Excel.quit()
    IF($_ -match "*.xlsx"){
        Write-host "This is a Excel Error"
        $Excel.quit()
        $ObjForm.Close()
        ER
$SCError =$_
$SCError2 = $SCError.exception.Message
$SCError | Out-File $ERL3 -append
$SCError2 | Out-File $ECL3 -append
#Write-Host "Script Ended Prematurely please see $ERL3" -ForegroundColor Magenta
$Res = "Script Ended Prematurely please see $ERL3"
$a = new-object -comobject wscript.shell
$b = $a.popup("$Res ",3," Error ",48)
#Write-Error $_
Return
    }
    $Excel.quit()
$ObjForm.Close()
ER
$SCError =$_
$SCError2 = $SCError.exception.Message
$SCError | Out-File $ERL3 -append
$SCError2 | Out-File $ECL3 -append
$Res = "Script Ended Prematurely please see $ERL3"
$a = new-object -comobject wscript.shell
$b = $a.popup("$Res ",3," Error ",48)
#Write-Host "Script Ended Prematurely please see $ERL3" -ForegroundColor Magenta
#Write-Error $_
Return}
}








