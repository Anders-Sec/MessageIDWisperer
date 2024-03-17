# MessageIDWisperer

Simple description

## Requirements
MessageIDWisperer requires the following PowerShell modules to function. The script will check if they are available and download them from the default PowerShell repository if they are not.
* ExchangeOnlineManagement

The following AzureAD role is required to run MessageIDWisperer
* Compliance Administrator

Unified Audit Logs need to be enabled and the user being investigated needs to have had a E5 License prior to the start of the compromise.

## Downloading MessageIDWisperer
You can download the latest version of MessageIDWisperer by using the following command:
```
git clone https://github.com/Anders-Sec/MessageIDWisperer.git
```

## Usage
The script is split up into three main functions: Identify, Review and, Export. The Identify function serves to pull all of the IPs used by that user during the given time window. This helps the analyst identify the malicious IPs used during the compromise. The Review function then takes in those IPs and generates a report of the total impact those IPs had. The last function Export, pulls down all the impacted mail items and exports them for further review by a legal team.

### First Time Setup
The script has a single config item that must be set before beginning an investigation. At the top of the script the $Domain variable must be set to the Analyst's domain (i.e. if the user has a UPN of User@domain.com, the Domain should be set to $Domain = "domain.com")

### Identify
The Identify function serves to report all the IPs used by the given User in a Time frame. This information can then be used to Identify any bad IPs for further review. The Identify functions requires only the target user for the investigation. optionally you can specify the time frame otherwise it will default to the previous 30 days. An example of the Identify function is given below:
```powershell
.\MessageIDWhisperer.ps1 -Function Identify -User victom@domain.com
```
This command will search for all IP used by the user "victom@domain.com" in the last 30 days. It will use the currently logged in user to authenticate to Exchange Online.

If you wish to authenticate as a different user than the currently logged in user you may use the -a [-Analyst] flag to specify a different UPN. An example of that is below:
```powershell
.\MessageIDWhisperer.ps1 -Function Identify -User victom@domain.com -StartDate 3/1/24 -EndDate 3/8/24 -Analyst analyst@domain.com
```

### Review
The Review function provides analysts a quick overview of the impact a compromised user might have. It will check the total number of emails accessed by the provided bad IPs. This functions requires the target user and a comma separated list of known bad IPs. An example of how to use the Review function is provided below:
```powershell
.\MessageIDWhisperer.ps1 -Function Review -User victom@domain.com -IP 8.8.8.8
```

While not required, the StartDate and EndDate flags can be used to narrow down the time period of the search. The Analyst flag can also be used in the Review function.
```powershell
.\MessageIDWhisperer.ps1 -Function Review -User victom@domain.com -StartDate 3/1/24 -EndDate 3/8/24 -Analyst analyst@domain.com -IP 8.8.8.8,1.1.1.1
```

### Export
The Export function will pull down as much information as possible about the impacted email items. The Export function utilizes the Get-MessageTrace cmdlet which can only pull details from emails with the last 10 days. If the email is from earlier than 10 days the script will still create a txt document with the name matching the Message ID but the contents will be empty. For these older emails, other tools like E-Discovery and the built-in Exchange function will be able to pull more information from these emails. Eventually the Start-HistoricalSearch cmdlet will be used to provide better search functionality for these older emails. The Export function requires the Target User and IP flags. An example of its usage is provided below:
```powershell
.\MessageIDWhisperer.ps1 -Function Export -User victom@domain.com -IP 8.8.8.8
```
By default, each Message ID is saved in a folder created in the scripts root directory named after the target user (ex. C:\MessageIDWisperer\victom@domain.com\MessageID.txt). This can be changed by providing the -o [-OutFolder] flag. An example of that is below:
```powershell
.\MessageIDWhisperer.ps1 -Function Export -User victom@domain.com -StartDate 3/1/24 -EndDate 3/8/24 -Analyst analyst@domain.com -IP 8.8.8.8 -OutFolder 'C:\Forensics\Case123\'
```


## Planed Future Features
* The Start-HistoricalSearch cmdlet will be used to provider better insight into older emails.
* The User and IP flags  will be able to parse in a file containing the listed items to search.
* There are some cases where the Unified Audit might be missing data, the script will check these conditions to verify that all information presented is all the data and no information is missing.
* An additional option in the Export function to reach out to the local exchange server and pull the message body for review by a legal team.
