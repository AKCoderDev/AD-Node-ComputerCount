# AD-Node-ComputerCount
PowerShell script for counting computer objects in Active Directory under specific nodes (OU or CN) by name.  Generates a human-readable TXT report with total counts and per-location summary.  Easily customizable by changing a single parameter ($TargetNodeName) to target different AD nodes.




Detailed Instructions for Using the Script
Purpose of the script:
This PowerShell script counts all computer objects located under Active Directory nodes (OU or CN) with a specific name. It then generates a human-readable TXT report that includes:

The total number of computers under all matching nodes.
A summary grouped by location (the OU directly above the target node).


1. What the script does

Searches your Active Directory for nodes named according to the variable TargetNodeName.
Counts all computer objects under each found node (including sub-OUs).
Extracts the location name from the Distinguished Name (DN) of each node.
Displays results in the console and saves them to a TXT file.


2. Before you start

Make sure RSAT (Remote Server Administration Tools) with the Active Directory module is installed.

On Windows 10/11:
Go to Settings → Apps → Optional Features → Add a feature → RSAT: Active Directory Domain Services and Lightweight Directory Tools.


Ensure your computer can connect to a Domain Controller (DC).
Verify DNS settings point to your domain DNS servers.
Confirm your system time is synchronized with the domain (Kerberos requires this).
You need at least read permissions in Active Directory.


3. How to prepare the script

Save the script locally, for example:
C:\scripts\AD-Node-ComputerCount-TXT-DN.ps1
Open the script in a text editor (PowerShell ISE or VS Code).
Edit the following parameters at the top of the script:

TargetNodeName → Set this to the name of the OU or CN you want to search.
Example:
$TargetNodeName = "ComputersAdministration"


OutputDir → Folder where the TXT report will be saved.
Example:
$OutputDir = "C:\scripts"


Server and CredentialStr → Optional. Use these if you need to connect to a specific DC or use different credentials.
Example:
$Server = "dc01.domain.local"
$CredentialStr = "DOMAIN\username"


IncludeCN → Set to $true if you also want to search CN containers named the same as TargetNodeName.




4. How to run the script

Open PowerShell as Administrator.
Navigate to the folder where the script is saved:
cd C:\scripts


Run the script:
powershell.exe -ExecutionPolicy Bypass -File .\AD-Node-ComputerCount-TXT-DN.ps1


If you see an error about execution policy, run:
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned




5. What happens when you run it

The script will:

Search for all OU (and optionally CN) named as TargetNodeName.
Count all computer objects under each node.
Display:

Total number of computers.
Summary by location.


Save a TXT report to the folder specified in OutputDir.



Example console output:
Found 23 target nodes named 'ComputersAdministration'
TOTAL computers under nodes 'ComputersAdministration': 157
Summary by location:
Location    Total
Office1      42
Office2       11
Office3       34

Example TXT report:
AD report for target node name: 'ComputersAdministration'
Generated: 2026-01-05 09:12:00

TOTAL computers under 'ComputersAdministration': 157

Summary by location (location = DN element directly above 'ComputersAdministration'):
Office1 : 42
Office2 : 11
Office3 : 34


6. How to customize

To count computers in a different OU name, change only:
$TargetNodeName = "NewOUName"


To include CN containers, set:
$IncludeCN = $true


To change the output folder:
$OutputDir = "D:\Reports"




7. Troubleshooting

Error: ActiveDirectory module not found
→ Install RSAT and restart PowerShell.
Error: The server is not operational
→ Check VPN, DNS, and connectivity to DC.
Empty total in report
→ Ensure the OU actually contains computer objects.
Execution policy blocks script
→ Use -ExecutionPolicy Bypass when running the script.
