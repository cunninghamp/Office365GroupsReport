﻿<#
.SYNOPSIS
Reports on Office 365 Groups that have been created or modified since the script last ran.

.DESCRIPTION 
This script provides an email report of Office 365 Group status. Groups that have been created, modified, or unchanged are shown in the report.

.OUTPUTS
Email to defined recipient(s).

.PARAMETER UseCredential
Credentials to pass to Connect-EXOnline

.EXAMPLE
.\Get-O365GroupReport.ps1

.LINK
https://github.com/cunninghamp/Office365GroupsReport

.NOTES

Copyright (c) 2016 Paul Cunningham

Permission is hereby granted, free of charge, to any person obtaining a copy 
of this software and associated documentation files (the "Software"), to deal 
in the Software without restriction, including without limitation the rights 
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell 
copies of the Software, and to permit persons to whom the Software is 
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all 
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR 
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE 
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER 
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING 
FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER 
DEALINGS IN THE SOFTWARE.

#>

[CmdletBinding()]
param (
        [Parameter(Mandatory=$false)]
        [string]$UseCredential
)


#...................................
# Variables
#...................................

$now = Get-Date

$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$XMLFileName = "$($myDir)\UnifiedGroups.xml"

$reportemailsubject = "Office 365 Groups Report"

$NewGroups = @()
$ModifiedGroups = @()
$DeletedGroups = @()
$UnmodifiedGroups = @()


#...................................
# Email Settings
#...................................

#TODO - Pull these from local XML instead of hard-coding in script
#TODO - Dynamically determine SmtpServer based on MX of customer's domain?
# If the $smtpSettings.SmtpServer value is either "" or $null, the script
# will attempt to automatically derive the SMTP server from the recipient
# domain's MX records
$smtpsettings = @{
	To =  "paul@practical365.com"
	From = "paul@practical365.com"
	Subject = "$reportemailsubject - $now"
	SmtpServer = $null
	}

#...................................
# Script
#...................................


#Check for previous results
if (Test-Path $XMLFileName) {
    
    #JSON file found, ingest as last results
    $LastResults = Import-Clixml -Path $XMLFileName
}
else {
    Write-Verbose "No previous results found."
}

#Check whether an EXO remote session is already established and requires cmdlet is available
try {
    Get-Command Get-UnifiedGroup -ErrorAction STOP | Out-Null
}
catch {
    Write-Verbose "Get-UnifiedGroup cmdlet not available. Need to connect to Exchange Online."

    try {
        Get-Command Connect-EXOnline -ErrorAction STOP | Out-Null
    }
    catch {
        Write-Warning $_.Exception.Message
        Write-Warning "I recommend adding Connect-EXOnline to your PowerShell profile."
        Write-Warning "Refer to: https://github.com/cunninghamp/Office-365-Scripts/tree/master/Connect-EXOnline"
        EXIT
    }

    #Check if a stored credential is available to be used
    if ($UseCredential) {
        try {
            Connect-EXOnline -UseCredential $UseCredential
        }
        catch {
            throw $_.Exception.Message
        }
    }
    else {
        Write-Warning "Admin credentials are required for connecting to Exchange Online."
        $Credential = Get-Credential -Message "Enter your Exchange Online administrative credentials."
        if ($Credential -ne $null) {
            $EXOSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $URL -Credential $Credentials -Authentication Basic -AllowRedirection -Name "Exchange Online"
            Import-PSSession $EXOSession
        }
        else {
            throw "A credential was not provided."
        }
    }

}


#...................................
# Retrieve current list of Groups
#...................................

$UnifiedGroups = @(Get-UnifiedGroup | Select Guid,DisplayName,AccessType,Notes,ManagedBy,WhenCreated,WhenChanged)

$LastResultsGuids = $LastResults.Guid

#Loop through Guids and check whether they exist in previous results
foreach ($Guid in $UnifiedGroups.Guid) {
    Write-Verbose "Checking if $Guid exists in last results"

    if ($LastResultsGuids -icontains $Guid) {
        
        #Build a custom object to store the group's details for the report and any changed properties
        $GroupObject = New-Object -TypeName PSObject

        $HasChanged = $false
        $Changes = @()

        $CurrentObject = $UnifiedGroups | Where {$_.Guid -eq $Guid}
        $PreviousObject = $LastResults | Where {$_.Guid -eq $Guid}

        $GroupProperties = @($CurrentObject | Get-Member | Where {$_.MemberType -eq "NoteProperty"}).Name

        Write-Verbose "Group ""$($CurrentObject.DisplayName)"" exists in previous results"

        #Compare each property to determine if any have changed
        foreach ($Property in $GroupProperties) {
            
            if ($CurrentObject.$Property -ieq $PreviousObject.$Property) {
            
                Write-Verbose "No change detected for $Property"
                $GroupObject | Add-Member NoteProperty -Name $Property -Value $CurrentObject.$Property
            }
            else {

                Write-Verbose "$Property is different (was $($PreviousObject.$Property), and is now $($CurrentObject.$Property))"
                $HasChanged = $true
                $GroupObject | Add-Member NoteProperty -Name $Property -Value "$($CurrentObject.$Property) (was $($PreviousObject.$Property))"
            }
        }

        #Add the group to either the modified or unmodified list
        if ($HasChanged) {
            
            $ModifiedGroups += $GroupObject

        }
        else {
        
            $UnmodifiedGroups += $GroupObject
        
        }
            
    }
    else {

        Write-Verbose "Group does not exist in previous results, therefore is a new group"

        $NewObject = $UnifiedGroups | Where {$_.Guid -eq $Guid}

        $NewGroups += $NewObject

    }
}


Write-Verbose "============ Summary ============"

Write-Verbose "New groups: $($NewGroups.Count)"
Write-Verbose "Modified groups: $($ModifiedGroups.Count)"
Write-Verbose "Unmodified groups: $($UnmodifiedGroups.Count)"

#Output current Groups info to XML for next run
#TODO - preserve last X copies of XML file as a backup for troubleshooting
try {
    Write-Verbose "Writing current groups info to XML for comparison on next run."
    $UnifiedGroups | Export-Clixml -Path $XMLFileName -ErrorAction STOP
}
catch {
    Write-Warning $_.Exception.Message
}

#...................................
# Validate SMTP Settings
#...................................

# If there's no SMTP Server specified, attempt to derive one from MX records
if ([string]::IsNullOrWhiteSpace($smtpSettings.SmtpServer)) {
    Write-Verbose "No SMTP server was specified - deriving one from DNS"
    try {
        $recipientSmtpDomain = $smtpsettings.To.Split("@")[1]
        $MX = Resolve-DnsName -Name $recipientSmtpDomain -Type MX | 
            Where-Object {$_.Type -eq "MX"} | 
            Sort-Object Preference | 
            Select-Object -First 1 -ExpandProperty NameExchange
        Write-Verbose "Found MX record: '$MX'"
        $SmtpSettings.SmtpServer = $MX
    } catch {
        throw "Unable to resolve SMTP Server and none was specified.`n$($_.Exception.Message)"
    }
}

#...................................
# Build the report
#...................................

#HTML HEAD with styles
$htmlhead="<html>
			<style>
			BODY{font-family: Arial; font-size: 10pt;}
			H1{font-size: 22px;}
			H2{font-size: 20px; padding-top: 10px;}
			H3{font-size: 16px; padding-top: 8px;}
			TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt; table-layout: fixed;}
            TABLE.testresults{width: 850px;}
            TABLE.summary{text-align: center; width: auto;}
			TH{border: 1px solid black; background: #dddddd; padding: 5px; color: #000000;}
            TH.summary{width: 80px;}
            TH.test{width: 120px;}
            TH.description{width: 150px;}
            TH.outcome{width: 50px}
            TH.comments{width: 120px;}
            TH.details{width: 270px;}
            TH.reference{width: 60px;}
			TD{border: 1px solid black; padding: 5px; vertical-align: top; }
			td.pass{background: #7FFF00;}
			td.warn{background: #FFE600;}
			td.fail{background: #FF0000; color: #ffffff;}
			td.info{background: #85D4FF;}
            ul{list-style: inside; padding-left: 0px;}
			</style>
			<body>"

#HTML intro
$IntroHtml="<h1>Office 365 Groups Report</h1>
			<p><strong>Generated:</strong> $now</p>"

#HTML report body

#TODO - Add handling for zero results here so report doesn't get mangled
$NewGroupsIntro = "<p>New Groups found:</p>"
$NewGroupsTable = $NewGroups | ConvertTo-Html -Fragment

#TODO - Add handling for zero results here so report doesn't get mangled
$ModifiedGroupsIntro = "<p>Modified Groups found:</p>"
$ModifiedGroupsTable = $ModifiedGroups | ConvertTo-Html -Fragment

#TODO - Add handling for zero results here so report doesn't get mangled
$UnmodifiedGroupsIntro = "<p>Unmodified Groups found:</p>"
$UnmodifiedGroupsTable = $UnmodifiedGroups | ConvertTo-Html -Fragment

$ReportBodyHtml = $NewGroupsIntro + $NewGroupsTable + $ModifiedGroupsIntro + $ModifiedGroupsTable + $UnmodifiedGroupsIntro + $UnmodifiedGroupsTable

#HTML TAIL
$htmltail = "<p>Report created by <a href=""http://practical365.com"">Practical365</a>.</p>
            </body>
			</html>"

$htmlreport = $htmlhead + $IntoHtml + $ReportBodyHtml + $htmltail

#TODO - Add option to output to HTML file

#TODO - Make this a parameter/switch
try {
    Send-MailMessage @smtpsettings -Body $htmlreport -BodyAsHtml -ErrorAction STOP
    Write-Verbose "Email report sent."
}
catch {
    Write-Warning $_.Exception.Message
    Write-Verbose "Email report not sent."
}

#...................................
# Finished
#...................................