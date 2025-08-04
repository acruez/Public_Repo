<#########################################################################
##########################################################################
 
    __  ___                        ______                       ____  _            
   /  |/  /___  ________  _____   / ____/___  ____  _______  __/ / /_(_)___  ____ _
  / /|_/ / __ \/ ___/ _ \/ ___/  / /   / __ \/ __ \/ ___/ / / / / __/ / __ \/ __ `/
 / /  / / /_/ (__  )  __/ /     / /___/ /_/ / / / (__  ) /_/ / / /_/ / / / / /_/ / 
/_/  /_/\____/____/\___/_/      \____/\____/_/ /_/____/\__,_/_/\__/_/_/ /_/\__, /  
                                                                          /____/   

##########################################################################
#########################################################################>
# Created 8/17/23 by Aaron Cruez.
# Last updated 1/31/2025 by Aaron Cruez.
# This script is used for seeing all emails sent from a Shared Mailbox. It is used to see which user with delegated permissions
# actually sent the email and details about the email.
##########################################################################
# Installs the ExchangeOnlineManagement module if its not alreay installed.
Install-Module ExchangeOnlineManagement -Scope AllUsers
# Imports the ExchangeOnlineManagment Module
Import-Module ExchangeOnlineManagement
# Connect to Exchange Online. You will get a browser popup to authenticate to. Once authenticated, you can close the browser that popped up. 
Connect-ExchangeOnline

#Checks if the C:\Temp folder exists, if not it creates it.
if( -Not ( Test-Path C:\Temp)){
    New-item -path C:\ -Name Temp -ItemType Directory
}
# Checks if the C:\Temp\SharedMailboxReport folder exits, if not it creates it.
if( -Not ( Test-Path C:\Temp\SharedMailboxReport)){
    New-item -path C:\Temp -Name SharedMailboxReport -ItemType Directory
}

# Asks the user to enter the email address of the shared mailbox they would like to search against. Makes them confirm that is the correct email address.
# Checks to make sure what the user entered is a vaild mailbox in the M365 Tenant. If not, it will loop and continue the process until they enter and confirm
# a valid email address.
do {
    $SharedMailbox = Read-Host "Enter the email address of the shared mailbox you would like to search against:"
    $SharedMailboxQuestionConfirmation = Read-Host "Is $SharedMailbox the correct email address? Enter 'Yes' to confirm:"
    $SharedMailboxChallenge = Get-Mailbox -Identity $SharedMailbox
} until ( $SharedMailboxQuestionConfirmation -eq 'yes' -and $SharedMailboxChallenge -ne $null)

# Asks the user to enter the number of days they would like to search against. The max range is 90 days. The script varifies the value is greater than 0 but
# less than 90. Asks the user to confirm that is the correct number. If not, it will loop and continue the process until they enter number that is valid for the
# search parameters.
do {
    do{ [Int64]$NumberOfDays = Read-Host "Enter the number of days you would like to search against (90 days is the maximum):"
    } until ($NumberOfDays -le 90 -and $NumberOfDays -gt 0)
    $NumberofDaysConfirmation = Read-Host "Is $NumberOfDays correct? Enter 'Yes' to confirm:"
} until ( $NumberofDaysConfirmation -eq 'yes')

# Creates the date range based on users input. The $DateRangeStartDate subtracts the number of days the user spcified from the present day the script is ran
$DateRangeStartDate = (Get-Date).AddDays(-$NumberofDays).ToString('MM-dd-yyyy')
# Creates the date range. The $DateRangeEndDate is the present day the script is ran.
$DateRangeEndDate = (Get-Date).ToString('MM-dd-yyyy')
# Creates a variable for the name the report will be given
$ReportName = ($SharedMailbox + " past $NumberofDays day search")

# Starts a MailboxAuditLog 
Search-MailboxAuditLog $SharedMailbox -LogonTypes Admin, Owner,Delegate -ShowDetails -StartDate $DateRangeStartDate| Export-CSV c:\temp\SharedMailboxReport\"Audit Log.CSV" –NoTypeInformation -Encoding UTF8
# Starts a HistoricalSearch with the date range specified.
Start-HistoricalSearch -ReportTitle $ReportName -StartDate $DateRangeStartDate -EndDate $DateRangeEndDate -ReportType MessageTrace -SenderAddress $SharedMailbox

# Gets information for all historical searches and stores in the variable $HistoricalSearch
$HistoricalSearch = @(Get-HistoricalSearch)
# Exports the results of the Historical Search to a csv file.
$HistoricalSearch | Export-CSV c:\temp\SharedMailboxReport\"HistoricalSearch.CSV" –NoTypeInformation -Encoding UTF8
# Imports the csv to the variable $ImportCSV
$ImportCSV = Import-Csv c:\temp\SharedMailboxReport\"HistoricalSearch.CSV"
# The $ReportNameCSV pulls the information relating to the report you created.
$ReportNameCSV = @($ImportCSv | Where-Object ReportTitle -eq $ReportName)
# Pulls the jobid for the report created
$Jobid = $ReportNameCSV.JobId

# The $StatusCheck gets the info for the report created by Job id
$StatusCheck = Get-HistoricalSearch -JobId $Jobid
# The $StatusCheck gets the job status for the report created. This can either be "Not Started", "In Progress", or "Done"
$StatusCheck.Status

# Starts a do/until loop to keep checking the status of the report until the status shows "Done". The script will start a sleep timer for 5 minutes
# and then check again. Once the status shows "Done", the script will exit the loop and continue with the rest of the script.
Do{
    $StatusCheck = Get-HistoricalSearch -JobId $Jobid
    Write-Host "The report status is: " $Statuscheck.status
    Write-Host "Script checks the status again every 5 minutes until the status shows 'Done'. Please wait this can take some time."
    Start-Sleep -Seconds 300
} until ($StatusCheck.Status -eq "Done")

# Gets an updated historical search
$HistoricalSearchDone = @(Get-HistoricalSearch)
# Exports out the updated historical search to csv
$HistoricalSearchDone | Export-CSV c:\temp\SharedMailboxReport\"HistoricalSearchDone.CSV" –NoTypeInformation -Encoding UTF8
# Imports the csv to a variable
$ImportCSVDone = Import-Csv c:\temp\SharedMailboxReport\"HistoricalSearchDone.CSV"

$ReportNameCSVDone = @($ImportCSvDone | Where-Object ReportTitle -eq $ReportName)
$ReportStatusDescription = $ReportNameCSVDone.reportstatusdescription

If ($ReportStatusDescription -ne "Complete - No results found"){
    $ReportFileUrl = $ReportNameCSVDone.fileurl
    Write-Host "Please use the url below in a web browser. It is a direct download for the report. If prompted enter your administrative m365 credentials:
    $ReportFileUrl" 
    do {
        $DownloadQuestion = Read-Host "Please confirm that the file downloaded to the Downloads folder. Enter 'Yes' once confirmed: "
    } until ( $DownloadQuestion -eq 'yes')

    $DownloadedFileName = "MTSummary_" + $ReportName + "_" + $Jobid + ".csv"
    Move-item -Path "$env:USERPROFILE\Downloads\$DownloadedFileName" -Destination "C:\temp\SharedMailboxReport\$DownloadedFileName"
    #
    #
    #
    #
    #
    #
    #
    #
    #
    #
}

Else {Write-host ("The report has completed but found no results in the last " + "$NumberOfDays" + " days.")}
####################################################################################################################################################
####################################################################################################################################################

$AduitLogCSV = Import-Csv 'C:\Temp\SharedMailboxReport\Audit Log.CSV' | Select-Object ItemInternetMessageID,LogonUserDisplayName,Operation,MailboxOwnerUPN,ItemSubject,LastAccessed
$AduitLogCSV | Export-CSV "C:\Temp\SharedMailboxReport\InfoMerge.CSV"
$AuditLogFilteredCSV = Import-CSV "C:\Temp\SharedMailboxReport\InfoMerge.CSV"
$DownloadedReportCSV = Import-CSV "C:\Temp\SharedMailboxReport\MTSummary_accountsreceivable@enterprisecompanytest.com past 9 day search_07b1fd4d-dd4b-4e25-8167-eb48489126f0.csv" -Encoding Unicode
$DownloadedReportCSV | Export-Csv "C:\Temp\SharedMailboxReport\SharedReport.csv"
$SharedReportCSV = Import-Csv "C:\Temp\SharedMailboxReport\SharedReport.csv" | Select-Object message_id,recipient_status

foreach ($Row in $AuditLogFilteredCSV){
#    write-host $row.ItemInternetMessageId
    If ($row.ItemInternetMessageId -eq $SharedReportCSV.message_id){
        $AuditLogFilteredCSV | Add-member -NotePropertyName "Recipients Email Address" -NotePropertyValue ("Test")
    }
    Else{Write-host "Test"}
}

$AuditLogFilteredCSV | Export-CSV C:\Temp\SharedMailboxReport\FinalReport.csv
##########################################################################
##########################################################################