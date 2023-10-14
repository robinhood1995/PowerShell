Function ExcelCSV ($File)
{

    #Modified by SFL 060118
    #Modified By SFL 061018
    #Modified by SFL 061318
    #modified by SFL 091018

    #https://stackoverflow.com/questions/7834656/create-log-file-in-powershell
    #https://www.sqlshack.com/connecting-powershell-to-sql-server/
    #This will look in the archive folder anowdat the last file generated anowdat check it againt the current file's Last Write DateTime
    #If the same it will email the team to have a look at why did we not get a new ExRate update

    $SQLServer = "hskiw-espdb-pg1\kiwprod"
    $SQLDatabase = "Stora_Live"
    $FileServer = "\\HSKIW-ESPFS-P01.group.corp.storaenso.com\Transferdata"
    $WorkFolder = "ExchangeRates"
    $BankName = "European Central Bank"
    $BankCode ="ECB"

    $toWhoSuccess = @("steve.ling@kiwiplan.com", "patrick.lorentz@kiwiplan.com", "elena.eskova@Storaenso.com","Maciej.Zajkowski@storaenso.com","Aleksanowdater.Dovgosheya@storaenso.com")
    $toWhoFail = @("steve.ling@kiwiplan.com", "patrick.lorentz@kiwiplan.com", "elena.eskova@Storaenso.com","Maciej.Zajkowski@storaenso.com","Aleksanowdater.Dovgosheya@storaenso.com")
    #$toWhoSuccess = @("steve.ling@kiwiplan.com")
    #$toWhoFail = @("steve.ling@kiwiplan.com")

# Do not change anything below this section#

    # Global Variables
    $nowdat = (get-date).ToString("yyyyMMddHHmm")
    $logDir = "$FileServer\$WorkFolder\Log"
    $logfile = "$logDir\$nowdat.$BankCode.log"
    $dayofweek = ( get-date ).DayOfWeek.value__
    $excludestarttime = "2300"
    $excludeendtime = "2359"

    # Excel file from the bank
    $excelFile = "$FileServer\ExchangeRates\$($BankCode)exchangerates.xlsx"
    $excellastdt = ([datetime](Get-ItemProperty -Path $excelFile -Name LastWriteTime).lastwritetime).ToString("yyyyMMddHHmm")
    
    #File Check Logs
    $logSuccess = "$logDir\$excellastdt.$BankCode.Success.log"

    # Convert csv file
    $csvFile = "$FileServer\$WorkFolder\$BankCode csv\archive\$($BankCode)exchangerates.csv_"

    # Last converted csv file
    $lastcsvdir = "$FileServer\$WorkFolder\$BankCode csv\archive"
    $lastcsvdtname = Get-ChildItem -Path $lastcsvdir | Sort-Object LastAccessTime -Descending | Select-Object -First 1
    $lastcsvdt = $lastcsvdtname.name    
    $lastcsvfilecheck = $csvFile+$excellastdt
    
    # ESP last processed file
    $ProcessedFileChk = "$FileServer\$WorkFolder\Processed\$($lastcsvdt)*"

    # Look into ESP database
    $sqlConn = New-Object System.Data.SqlClient.SqlConnection
    $sqlConn.ConnectionString = “Server=$SQLServer;Integrated Security=true;Initial Catalog=$SQLDatabase”
    $sqlConn.Open()
    #$sqlcmd = $sqlConn.CreateCommand()
    $sqlcmd = New-Object System.Data.SqlClient.SqlCommand
    $sqlcmd.Connection = $sqlConn
    $query = “select count(*) from orgexchangeRate where cast(effectivedate as date) = cast(getdate() as date) or cast(mainttime as date) = cast(getdate() as date)”
    $sqlcmd.CommandText = $query
    $count = 0
    $count = $sqlcmd.ExecuteScalar()
    
    Write-Log "INFO" "$BankName" $logfile
    
    #Exclude for Sunday 0 anowdat Monday 1
    if ( ($dayofweek -eq 0) -or ($dayofweek -eq 1) ) {
    Write-Log "INFO" "It is the Weekend Nothing to do" $logfile
    Write-Log "WARN" "Exiting the process" $logfile
    exit
    }
    else {
    Write-Log "INFO" "It's not the Weekend keep going" $logfile
    }

    #exculde running between these times
    if ( ($excludestarttime -ge (get-date -Format "HHmm")) -and ($excludeendtime -le (get-date -Format "HHmm")) ) {
    Write-Log "INFO" "The time is between $excludestarttime and $excludeendtime and we are not processing at this time" $logfile
    Write-Log "INFO" "The process retreives the new rate from the bank at 23:00" $logfile
    Write-Log "WARN" "Exiting the process" $logfile
    exit
    }
    else {
    Write-Log "INFO" "It's time to process some files" $logfile
    }

    Write-Log "INFO" "This is the $SQLDatabase database we are connected to on server $SQLServer" $logfile
    Write-Log "INFO" "This is excel file $excelFile retrieves data from the bank" $logfile
    Write-Log "INFO" "Date and time $excellastdt from the last live file that exists from the bank" $logfile
    Write-Log "INFO" "Verify if a the exchange rate file celled $($BankCode)exchangerates.csv_$excellastdt is not the same as the last processed file called $lastcsvdt" $logfile

    #Check the folder for a Sucess file and if so stop processing
    if ( ([System.IO.File]::Exists($logSuccess)) ) {
    Write-Log "INFO" "A success file exists $logSuccess" $logfile
    Write-Log "INFO" "Nothing more to do" $logfile
    Write-Log "WARN" "Exiting the process" $logfile
    exit
    } else {
    Write-Log "INFO" "No Success file was found called $logSuccess" $logfile
    }

    #Let's look at the current bank file and what also at the last file we converted to CSV
    if ( ($count -gt 0) -and ("$($BankCode)exchangerates.csv_$excellastdt" -eq $lastcsvdt) ) {
    
    Write-Log "INFO" "The two file are a match so lest look at the rates we imported" $logfile

    $query = “select mainttime, effectivedate, rate from orgexchangeRate where cast(effectivedate as date) = cast(getdate() as date) or cast(mainttime as date) = cast(getdate() as date)”
    $sqlcmd.CommandText = $query
    #This is if you wish to display the result
    $adp = New-Object System.Data.SqlClient.SqlDataAdapter $sqlcmd
    $data = New-Object System.Data.DataSet
    $adp.Fill($data) | Out-Null
    
    foreach ($Row in $data.Tables[0].Rows)
    { 
    Write-Log "INFO" "Import Time of $($Row.mainttime) Effective Date $($Row.effectivedate) Rate of $($Row.rate)" $logfile
    Write-Log "INFO" "Import Time of $($Row.mainttime) Effective Date $($Row.effectivedate) Rate of $($Row.rate)" $logSuccess
    }
    Write-Log "INFO" "ESP database has $count exchange rates imported today" $logfile
    Write-Log "INFO" "ESP database has $count exchange rates imported today" $logSuccess
    Write-Log "WARN" "Exiting the process" $logfile
    Send-MailMessage -To $toWhoSuccess -From "agent@storaenso.com" -Subject "INFO $WorkFolder $BankCode - $count ER imported into ESP" -SmtpServer "relay.storaenso.com" -Body "Log Directory $logDir" -Attachments "$logfile"
    exit
    } else {
    Write-Log "INFO" "Either the are no rates imported and the files match or the files do not match and we have no imported rates" $logfile
    }
        

    # Let add a new file as the file is there and ESP has no rates
    if ( ($count -eq 0) -or ("$($BankCode)exchangerates.csv_$excellastdt" -ne $lastcsvdt) ) {
    
    Write-Log "INFO" "The two file are not a match so let's add a file for ESP to import" $logfile
    Write-Log "INFO" "Adding the new file" $logfile
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false
    $wb = $Excel.Workbooks.Open($excelFile)
    #foreach ($ws in $wb.Worksheets)
    $wb.Sheets.Item("Sheet1")
    #{
        $wb.SaveAs("$FileServer\$WorkFolder\$($BankCode) csv\$($BankCode)exchangerates.csv_$excellastdt", 6)
        Write-Log "INFO" "Added file for Agent to pick up: $($BankCode)exchangerates.csv_$excellastdt" $logfile
        $wb.SaveAs("$FileServer\$WorkFolder\$($BankCode) csv\archive\$($BankCode)exchangerates.csv_$excellastdt", 6)
        Write-Log "INFO" "Added file for Archive: $($BankCode)exchangerates.csv_$excellastdt" $logfile
            
    #}
    Write-Log "INFO" "A new file was created for ESP to process" $logfile
    Send-MailMessage -To $toWhoSuccess -From "agent@storaenso.com" -Subject "$WorkFolder $BankCode New ESP import file created to import" -SmtpServer "relay.storaenso.com" -Body "Log Directory $logDir" -Attachments "$logfile"
    $Excel.Quit()
    
    }

    #Delete Old Log Files
    $oldlogfilesdays = 15
    Get-ChildItem $logDir -Recurse | Where-Object {-not $_.PSIsContainer -and (Get-Date).Subtract($_.CreationTime).Days -gt $oldlogfilesdays } | Remove-Item -WhatIf $logfile

}

Function Write-Log {
    [CmdletBinding()]
    Param(
    [Parameter(Mandatory=$False)]
    [ValidateSet("INFO","WARN","ERROR","FATAL","DEBUG")]
    [String]
    $Level = "INFO",

    [Parameter(Mandatory=$True)]
    [string]
    $Message,

    [Parameter(Mandatory=$False)]
    [string]
    $logfile
    )
    
    $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
    $Line = "$Stamp $Level $Message"
    #If($logfile) {
        Add-Content $logfile -Value $Line
    #}
    #Else {
        Write-Output $Line
    #}
}

$FileName = "Book1"
#Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
ExcelCSV -File "$FileName"