# Office 365 Group Report Script

Get-O365GroupReport.ps1 is a PowerShell script that generates an email report about new, modified, and deleted Office 365 Groups (also known as Unified Groups).

The script will store information about the Office 365 Groups in your tenant in a file named *UnifiedGroups.xml*, located in the same folder as the script. The first time you run the script, all Groups will be reported as "New". On subsequent runs, the script will use the previous results to determine which Groups are new, modified, or deleted since the last time the script was run.

## Installation

This script can make use of the following functions:

- [Get-StoredCredential](http://practical365.com/blog/saving-credentials-for-office-365-powershell-scripts-and-scheduled-tasks/)
- [Connect-EXOnline](https://github.com/cunninghamp/Office-365-Scripts/tree/master/Connect-EXOnline)

By adding those functions to your PowerShell profile you can streamline the running of Get-O365GroupReport.ps1. Without those functions, you'll need to run the script from an existing PowerShell session that is connected to Exchange Online, or you'll be prompted for credentials to connect to Exchange Online.

To install the script:

1. Download the latest release from [GitHub](https://github.com/cunninghamp/Office365GroupsReport/releases) or the [TechNet Script Gallery](https://gallery.technet.microsoft.com/office/Office-365-Groups-Report-7e3e161b).
2. Unzip the files to a folder on the workstation or server where you want to run the script.
3. Rename *Get-O365GroupReport.xml.sample* to *Get-O365GroupReport.xml*
4. Edit *Get-O365GroupReport.xml* with appropriate email settings for your environment. If you exclude the SMTP server, the script will send the report email to the first MX record for the domain of the *To* address.
5. Run the script using the usage examples below.

## Configuration

You can customize the behavior of the script by editing settings in the Get-O365GroupReport.xml file.

- **HistoryItemsToKeep:** defines the number of previous UnifiedGroups.xml files that will be preserved in the \history folder. The default is 10. History files might be useful to you for manual comparisons.
- **EmailOnlyIfChanges:** controls whether to send the email report if there are no new, modified, or deleted groups detected. The default is 1, which means the email report will not send if there's no changes. If you prefer to receive the report every time, change the value to 0. 

## Usage  

Run the script in a PowerShell console.

```
.\Get-O365GroupReport.ps1
```

Run the script, using a stored credential to connect to Exchange Online.

```
.\Get-O365GroupReport.ps1 -UseCredential admin@tenantname.onmicrosoft.com
```

Run the script with verbose output.

```
.\Get-O365GroupReport.ps1 -Verbose
```

Run the script with manual SMTP settings that override the Get-O365GroupReport.xml configuration.

```
.\Get-O365GroupReport.ps1 -MailFrom reports@contoso.com -MailTo you@contoso.com -MailSubject "Your custom subject" -SmtpServer mail.contoso.com
```

## Credits

Written by:

- Paul Cunningham, [Blog](http://practical365.com) | [GitHub](https://github.com/cunninghamp) | [Twitter](https://twitter.com/paulcunningham)
- Chris Brown, [Blog](https://www.flamingkeys.com) | [GitHub](https://github.com/chrisbrownie) | [Twitter](https://twitter.com/chrisbrownie)

For more Office 365 tips, tutorials, and news check out [Practical 365](http://practical365.com).
