<#  
.SYNOPSIS  
	This script reports failures in multiple Health Check information on Exchange 2016 servers

.PARAMETER ServerFilter
Default MSGSVR
The name or partial name of the Servers to query valid values such as MSGSVR or MSGCONSVR2301

.PARAMETER ReportPath
Default Current Folder
Directory location to create the HTML report

.PARAMETER IncludeCSV
Default False
Add CSV output file in ReportPath location

.PARAMETER Threads
Default 30
Number of simultaneous threads querying servers

.NOTES  
  Version      				: 0.2
  Rights Required			: Exchange View Only Admin/Local Server Administrator
  Exchange Version			: 2016/2013 (last tested on Exchange 2016 CU14/Windows 2012R2)
  Authors       			: Steven Snider (stevesn@microsoft.com) (additional html reporting code borrowed from internet examples)
  Last Update               : Nov 12 2019

.VERSION
  0.1 - Initial Version for connecting Internal Exchange Servers
  0.2 - Updating output formatting, colors, and strings for easier readability	
#>

Param(
   [Parameter(Mandatory=$false)] [string] $ServerFilter="MSGSVR",
   [Parameter(Mandatory=$false)] [string] $Threads=30,
   [Parameter(Mandatory=$false)] [boolean] $IncludeCSV=$False,
   [Parameter(Mandatory=$false)] [string] $ReportPath=(Convert-Path .)

)

#region Verifying Administrator Elevation
Write-Host Verifying User permissions... -ForegroundColor Yellow
Start-Sleep -Seconds 2
#Verify if the Script is running under Admin privileges
If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
  [Security.Principal.WindowsBuiltInRole] "Administrator")) 
{
  Write-Warning "You do not have Administrator rights to run this script.`nPlease re-run this script as an Administrator!"
  Write-Host 
  Break
}
#endregion

#If (-Not($UserCredential)) {
#    $UserCredential = Get-Credential
#}

[string]$search = "(&(objectcategory=computer)(cn=$serverfilter*))"
$ExchangeServers = ([adsisearcher]$search).findall() | %{$_.properties.name} | sort

$Servers = @()
$Servers = $ExchangeServers

#diagnostic block
### Change
#$Servers="SQL","Ex2016a","SP","Ex2016b","MIM"
#$Servers="Ex2016a"
#$IncludeCSV=$True

#region Script Information

Write-Host "--------------------------------------------------------------" -BackgroundColor DarkGreen
Write-Host "Exchange Health Report" -ForegroundColor Green
Write-Host "Version: 0.2" -ForegroundColor Green
Write-Host "--------------------------------------------------------------" -BackgroundColor DarkGreen
#endregion

$FileDate = "{0:yyyy_MM_dd-HH_mm_ss}" -f (get-date)
$ServicesFileName = $ReportPath+"\ExHealthReport-"+$FileDate+".html"
[Void](New-Item -ItemType file $ServicesFileName -Force)

If ($IncludeCSV -eq $True) {
   $ServicesCSVFileName = $ReportPath+"\ExHealthReport-"+$FileDate+".CSV"
   [Void](New-Item -ItemType file $ServicesCSVFileName -Force)

}



If ($Servers.count -eq 0) {
    Write-Host "Filter returned zero servers.  Please adjust filter and try again." -ForegroundColor Red
    Exit
}

If ($Servers.count -lt $Threads) {
    Write-Host "List of servers is less than the number of threads assigned ($Threads), lowering background threads to match server count." -ForegroundColor Red
    $Threads = $Servers.count
} 

#### Building HTML File ####
Function writeHtmlHeader
{
    param($fileName)
    $date = ( get-date ).ToString('MM/dd/yyyy')
    Add-Content $fileName "<html>"
    Add-Content $fileName "<head>"
    Add-Content $fileName "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>"
    Add-Content $fileName '<title>Exchange Health Report</title>'
    add-content $fileName '<STYLE TYPE="text/css">'
    add-content $fileName  "<!--"
    add-content $fileName  "td {"
    add-content $fileName  "font-family: Segoe UI;"
    add-content $fileName  "font-size: 11px;"
    add-content $fileName  "border-top: 1px solid #1E90FF;"
    add-content $fileName  "border-right: 1px solid #1E90FF;"
    add-content $fileName  "border-bottom: 1px solid #1E90FF;"
    add-content $fileName  "border-left: 1px solid #1E90FF;"
    add-content $fileName  "padding-top: 0px;"
    add-content $fileName  "padding-right: 0px;"
    add-content $fileName  "padding-bottom: 0px;"
    add-content $fileName  "padding-left: 0px;"
    add-content $fileName  "}"
    add-content $fileName  "body {"
    add-content $fileName  "margin-left: 5px;"
    add-content $fileName  "margin-top: 5px;"
    add-content $fileName  "margin-right: 0px;"
    add-content $fileName  "margin-bottom: 10px;"
    add-content $fileName  ""
    add-content $fileName  "table {"
    add-content $fileName  "border: thin solid #000000;"
    add-content $fileName  "}"
    add-content $fileName  "-->"
    add-content $fileName  "</style>"
    add-content $fileName  "</head>"
    add-content $fileName  "<body>"
    add-content $fileName  "<table width='100%'>"
    add-content $fileName  "<tr bgcolor='#336699 '>"
    add-content $fileName  "<td colspan='7' height='25' align='center'>"
    add-content $fileName  "<font face='Segoe UI' color='#FFFFFF' size='4'>Exchange Health Report - $date</font>"
    add-content $fileName  "</td>"
    add-content $fileName  "</tr>"
    add-content $fileName  "</table>"
}

Function writeTableHeader
{
    param($fileName)
    Add-Content $fileName "<tr bgcolor=#0099CC>"
    Add-Content $fileName "<td width='10%' align='center'><font color=#FFFFFF>Server</font></td>"
    Add-Content $fileName "<td width='15%' align='center'><font color=#FFFFFF>Status</font></td>"
    Add-Content $fileName "<td width='15%' align='center'><font color=#FFFFFF>Services Not Running</font></td>"
    Add-Content $fileName "<td width='10%' align='center'><font color=#FFFFFF>Health Set</font></td>"
    Add-Content $fileName "<td width='15%' align='center'><font color=#FFFFFF>Health Monitor</font></td>"
    Add-Content $fileName "<td width='15%' align='center'><font color=#FFFFFF>Target Resource</font></td>"
    Add-Content $fileName "<td width='10%' align='center'><font color=#FFFFFF>Component</font></td>"
    Add-Content $fileName "<td width='10%' align='center'><font color=#FFFFFF>State</font></td>"
    Add-Content $fileName "</tr>"
}

Function writeHtmlFooter
{
    param($fileName)
    Add-Content $fileName "</body>"
    Add-Content $fileName "</html>"
}

Function writeServiceInfo
{
    param($filename,$servername,$status,$servicesnotrunning,$healthset,$healthmonitor,$healthtarget,$component,$state)

    
     Add-Content $fileName "<tr>"
     Add-Content $fileName "<td align='center'>$servername</td>"
     If ($status -eq "Error connecting to server") {Add-Content $fileName "<td BGColor='#FF0000'>$Status</td>"}
     ElseIf ($status -eq "UnHealthy") {Add-Content $fileName "<td BGColor='#FFFF00'>$Status</td>"}
     ElseIf ($status -eq "Maintenance Mode Enabled") {Add-Content $fileName "<td BGColor='#00FF7F'>$Status</td>"}
     Else {Add-Content $fileName "<td>$Status</td>"}
     Add-Content $fileName "<td>$ServicesNotRunning</td>"
     Add-Content $fileName "<td align='center'>$HealthSet</td>"
     Add-Content $fileName "<td align='center'>$HealthMonitor</td>"
     Add-Content $fileName "<td>$HealthTarget</td>"
     Add-Content $fileName "<td>$Component</td>"
     Add-Content $fileName "<td>$State</td>"
}

Function sendEmail
    { param($from,$to,$subject,$smtphost,$htmlFileName)
        $body = Get-Content $htmlFileName
        $smtp= New-Object System.Net.Mail.SmtpClient $smtphost
        $msg = New-Object System.Net.Mail.MailMessage $from, $to, $subject, $body
        $msg.isBodyhtml = $true
        $smtp.send($msg)
    }

Function jobFailed
    { param ($Server,$ReportData)

          Write-Host "Error Connecting to server " $Server ", Please verify connectivity and permissions" -ForegroundColor Red

          $ReportData += New-Object pSObject -Property @{
             'Server' = $Server
             'Status' = "Error connecting to server" 
             'ServicesNotRunning' = "Unknown"
             'HealthSetName' = "Unknown"
             'HealthMonitor' = "Unknown"
             'HealthTargetResource' = "Unknown"
             'Component' = "Unknown"
             'State' = "Unknown"
          }

    }


########################### Main Script ###################################
writeHtmlHeader $ServicesFileName

        Add-Content $ServicesFileName "<table width='100%'><tbody>"
        Add-Content $ServicesFileName "<tr bgcolor='#0099CC'>"
        Add-Content $ServicesFileName "</tr>"

        WriteTableHeader $ServicesFileName


$s = New-PSSession -Name HealthReport –ConfigurationName Microsoft.exchange –ConnectionUri http://exchange.contoso.com/powershell -Authentication Kerberos
Import-PSSession $s -CommandName Get-DatabaseAvailabilityGroup -AllowClobber | Out-Null

#### Find which servers are in Maintenance Mode to reduce checks that still need to be made and populate appropriate server flags in array
Write-Host "Checking for servers currently in Maintenance Mode"

## TODO filter DAGs if your filter doesn't include them.  i.e. if I'm looking for DAG21 servers, I dont care about all the other servers in the environment if they are in MM

$MaintMode=@()
$ReportData=@()

$MaintMode=get-DatabaseAvailabilityGroup -status * | %{$_.ServersInMaintenance} | sort

Write-Output "The following servers are in MM: " $MaintMode

ForEach ($Server in $MaintMode) {
    $ReportData += New-Object pSObject -Property @{
        'Server' = $Server
        'Status' = "Maintenance Mode Enabled" 
        'ServicesNotRunning' = ""
        'HealthSetName' = ""
        'HealthMonitor' = ""
        'HealthTargetResource' = ""
        'Component' = ""
        'State' = ""
    }
}

$Servers = $Servers | ? {$_ -notin $MaintMode} | Sort

$SB = {
    param($server)

    $ServiceHealth=@()
    Test-ServiceHealth -Server $server -ErrorAction STOP | ? {$_.RequiredServicesRunning -eq 0} | % {$ServiceHealth += $_.ServicesNotRunning}
    $ServiceHealth=$ServiceHealth | Sort | Get-Unique
    If ($Servicehealth.Count -ne 0) {
    #Foreach ($Service in $ServiceHealth) {
       $ReportData += New-Object pSObject -Property @{
          'Server' = $Server
          'Status' = "Services Not Running" 
          'ServicesNotRunning' = $ServiceHealth
          'HealthSetName' = ""
          'HealthMonitor' = ""
          'HealthTargetResource' = ""
          'Component' = ""   
          'State' = ""
       }
    }


    # If if there are issues with services, skip checking the Server Health report

    If ($ServiceHealth.Count -eq 0) {

       $ServerHealth=@();Get-ServerHealth $server -ErrorAction STOP | ? {$_.AlertValue -eq "Unhealthy"} | Select Server,HealthSetName,Name,TargetResource,AlertValue
       If ($ServerHealth.Count -gt 0) {
           $ReportData += New-Object pSObject -Property @{
              'Server' = $ServerHealth.Server 
              'Status' = $ServerHealth.AlertValue 
              'ServicesNotRunning' = ""
              'HealthSetName' = $ServerHealth.HealthSetName
              'HealthMonitor' = $ServerHealth.Name
              'HealthTargetResource' = $ServerHealth.TargetResource
              'Component' = ""
              'State' = ""
          }
       }
    }

    # If if there are issues with services skip checking the component health

    If ($ServiceHealth.Count -eq 0) {

       $ComponentHealth=@();Get-ServerComponentState $server -ErrorAction STOP | ? {$_.State -ne "Active"} | Select ServerFQDN,Component,State
       If ($ComponentHealth.Count -gt 0) {

           $ReportData += New-Object pSObject -Property @{
              'Server' = $ComponentHealth.ServerFQDN
              'Status' = ""
              'ServicesNotRunning' = ""
              'HealthSetName' = ""
              'HealthMonitor' = ""
              'HealthTargetResource' = ""
              'Component' = $ComponentHealth.Component
              'State' = $ComponentHealth.State
           }
       }
    }
    #If there is anything to return, output it such that the Job can pick it up and pass it to the global data array
    If ($ReportData.count -gt 0) {$ReportData}
} #End Scriptblock


$InitSB = {

   Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn

}


$Started = 0 #Job counter
## TODO:  Use a progress bar instead of counter, look for % waiting, % complete, etc..

foreach ($Server in $Servers)
{       
   $JobCount = Get-Job -Name Health* | ? {$_.State -eq "Running"}
   If ($JobCount.Count -lt $Threads) {     
       #compare running jobs to defined threads and start a new one unless full 
       try
       {
          $Started++
          Start-Job -Name "Health$Server" -ScriptBlock $SB -ArgumentList $Server -InitializationScript $InitSB | Out-Null
          Write-Host "Started Job " $Started " of " $Servers.Count " for server " $Server -ForegroundColor Green
       } # end Try
       catch
       {
#          Write-Host "Error Connecting to server " $Server ", Please verify connectivity and permissions" -ForegroundColor Red

#          jobFailed $Server,$ReportData

          $ReportData += New-Object pSObject -Property @{
             'Server' = $Server
             'Status' = "Error connecting to server" 
             'ServicesNotRunning' = "Unknown"
             'HealthSetName' = "Unknown"
             'HealthMonitor' = "Unknown"
             'HealthTargetResource' = "Unknown"
             'Component' = "Unknown"
             'State' = "Unknown"
          }

          Continue
       } #end catch
        
   }  else {
      # check the jobs & wait until one finishes
      $JobCount | Wait-Job -Any -Timeout 120 | Out-Null
      # We have an open spot, start the next thread


      Start-Job -Name "Health$Server" -ScriptBlock $SB -ArgumentList $Server -InitializationScript $InitSB | Out-Null
      $Started++
      Write-Host "Started Job " $Started " of " $Servers.Count " for server " $Server -ForegroundColor Green

   }        
}


   Get-Job | Wait-Job -Timeout 240 | Out-Null #default value 120 changed JDM because APAC wasn't returning and showed bad result when servers were good. 

   $Jobs = Get-Job Health*


ForEach ($Job in $Jobs) {
   If ($Job.State -ne "Completed") {

        Write-Host "Error with process " $Job.Name ", Please verify connectivity and permissions" -ForegroundColor Red
        $FailedServer = $job.Name.Replace("Health","")
 
#        jobFailed $FailedServer,$ReportData


        $ReportData += New-Object pSObject -Property @{
           'Server' = $FailedServer
           'Status' = "Error connecting to server" 
           'ServicesNotRunning' = "Unknown"
           'HealthSetName' = "Unknown"
           'HealthMonitor' = "Unknown"
           'HealthTargetResource' = "Unknown"
           'Component' = "Unknown"
           'State' = "Unknown"
        }
   } Else {
  

        $Server = $job.Name.Replace("Health","")
        $ReportData += New-Object pSObject -Property @{
           'Server' = $Server
           'Status' = "Health Check Completed" 
           'ServicesNotRunning' = ""
           'HealthSetName' = ""
           'HealthMonitor' = ""
           'HealthTargetResource' = ""
           'Component' = ""
           'State' = ""
        }        
   }

}

   $ExHealth=Get-job Health* | Receive-Job

   Foreach ($ex in $exhealth) {

   Write-Host $ex 

   #Translate return values to match Data Array requirements
   [String]$status = ""
   [String]$servername = ""
   [String]$avc=$ex.alertvalue
   If ($ex.Server.length -gt 0) {$servername = $ex.Server}
   If ($ex.ServerFQDN.length -gt 0) {$servername = ($ex.serverfqdn.split("."))[0]}
   If ($avc -gt 0) {$status = $avc}
   If ($ex.status.length -gt 0) {$status = $ex.status}
   [String]$state = $ex.state



        $ReportData += New-Object pSObject -Property @{
           'Server' = $servername
           'Status' = $status
           'ServicesNotRunning' = $ex.ServicesNotRunning
           'HealthSetName' = $ex.HealthSetName
           'HealthMonitor' = $ex.Name
           'HealthTargetResource' = $ex.TargetResource
           'Component' = $ex.Component
           'State' = $state
        }        

   }
   ##ToDo Add sort by DAG here

   $ReportDataFinal = $ReportData | Sort Server

   foreach ($item in $ReportDataFinal)
      {
         writeServiceInfo $ServicesFileName $item.Server $item.Status $item.ServicesNotRunning $item.HealthSetName $item.HealthMonitor $item.HealthTargetResource $item.Component $item.state
      }
       
   Add-Content $ServicesFileName "</table>"

writeHtmlFooter $ServicesFileName

### Configuring Email Parameters
#sendEmail from@domain.com to@domain.com "Health State Report - $Date" SMTPS_SERVER $ServicesFileName

#Closing HTML
writeHtmlFooter $ServicesFileName

If ($IncludeCSV -eq $True) {
   $ReportDataFinal | Export-Csv -Path $ServicesCSVFileName -NoTypeInformation
}

Write-Host "`n`nThe File was generated at the following location: $ServicesFileName `n`nOpenning file..." -ForegroundColor Cyan
Invoke-Item $ServicesFileName


Get-Job Health* | Remove-Job
Get-PSSession | Remove-PSSession