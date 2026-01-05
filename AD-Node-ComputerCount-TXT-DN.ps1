<#
  AD-Node-ComputerCount-TXT-DN.ps1
  Purpose:
    Count computers located under **specific AD nodes** (OU or CN) named `TargetNodeName`
    and produce a human-readable TXT report with:
      - total count
      - per-location summary (location = DN element directly above the target node)

  How to use:
    1) Set $TargetNodeName to the AD node name you want to count (e.g., "ADM", "ComputersAdministration").
    2) (Optional) Set $Server and $CredentialStr if you need to connect to a specific DC or use specific credentials.
    3) Run the script. The report is saved to $TxtSummary.

  Notes:
    - The script extracts "Location" from DN using a regex: the OU directly above the target node.
      Example DN: OU=ADM,OU=Office1,OU=Sites,DC=example,DC=com  -> Location = Office1
    - Uses defensive coding:
      * Forces arrays via @( ... ) so .Count is always an integer.
      * Only counts objects with class 'computer' via -LDAPFilter.
    - By default searches OU (organizational units). You can include CN containers by toggling $IncludeCN.

  Author: AKCoderDev
  Version: 1.1
#>

# ==========================
# USER PARAMETERS (EDIT HERE)
# ==========================
# ---> Change ONLY this name to reuse the script for different AD nodes:
# e.g., "ADM", "ComputersAdministration", "Workstations", "Kiosks", etc.
$TargetNodeName = "ADM"   # EXAMPLE VALUE: replace with "ComputersAdministration" when needed

# Output paths (where the TXT report will be saved)
$OutputDir  = "C:\scripts"
$TxtSummary = Join-Path $OutputDir "AD_${TargetNodeName}_Summary.txt"

# Directory services connection (optional)
# If your laptop is in the domain and you have sufficient rights, you can leave these empty.
$Server        = "domain.local"                 # FQDN of the domain or a specific DC (e.g., dc01.domena.local). Leave empty if not needed.
$CredentialStr = "Domain\username"  # Domain\username. Leave empty to use current identity.

# Search options
# Set to $true if you ALSO want to search CN containers named $TargetNodeName (besides OU).
# If your environment uses only OU, keep it $false (slightly faster).
$IncludeCN = $false

# ==========================
# PREPARATION / ENVIRONMENT
# ==========================
try {
    # Import RSAT ActiveDirectory module; required for Get-AD* cmdlets
    Import-Module ActiveDirectory -ErrorAction Stop
} catch {
    Write-Error "ActiveDirectory module not available. Install RSAT: Active Directory and try again."
    return
}

# Ensure output directory exists
if (-not (Test-Path $OutputDir)) {
    try { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }
    catch {
        Write-Error "Cannot create output directory '$OutputDir'. Run PowerShell as Administrator or change path."
        return
    }
}

# Build common parameters for Get-AD* cmdlets based on optional Server/Credential
$adParams = @{}
if ($Server -and $Server.Trim()) { $adParams.Server = $Server }
if ($CredentialStr -and $CredentialStr.Trim()) {
    try {
        $Cred = Get-Credential -UserName $CredentialStr -Message "Enter password for $CredentialStr"
        $adParams.Credential = $Cred
    } catch {
        Write-Error "Failed to acquire credentials. $_"
        return
    }
}

# ==========================
# HELPER FUNCTIONS
# ==========================

# Extracts "Location" from DN as the element directly above the target node name.
# Example:
#   DN: OU=ADM,OU=Office1,OU=Sites,DC=example,DC=com
#   TargetNodeName: "ADM"
#   Location => "Office1"
function Get-LocationFromDN([string]$dn, [string]$nodeName) {
    if ([string]::IsNullOrWhiteSpace($dn) -or [string]::IsNullOrWhiteSpace($nodeName)) { return "UNKNOWN" }
    # Regex explanation:
    #   - Look for "OU=<nodeName>,OU=<capture>,"
    #   - The capture group 1 is the location (the OU just above the node)
    # If your target is a CN container, this still works (the search base is OU by default),
    # but you can modify the regex to 'CN=<nodeName>,OU=([^,]+),' if needed.
    $pattern = "OU=$([regex]::Escape($nodeName)),OU=([^,]+),"
    $m = [regex]::Match($dn, $pattern, 'IgnoreCase')
    if ($m.Success) { return $m.Groups[1].Value }

    # Fallback: try CN pattern (in case the target is under CN)
    $patternCN = "CN=$([regex]::Escape($nodeName)),OU=([^,]+),"
    $m2 = [regex]::Match($dn, $patternCN, 'IgnoreCase')
    if ($m2.Success) { return $m2.Groups[1].Value }

    return "UNKNOWN"
}

# ==========================
# DISCOVERY: FIND TARGET NODES
# ==========================
Write-Host "Searching AD for nodes named '$TargetNodeName'..." -ForegroundColor Cyan

$targets = @()

# 1) Find OU nodes named exactly as TargetNodeName
try {
    $ouFilter = "Name -eq '$TargetNodeName'"
    $ous = Get-ADOrganizationalUnit @adParams -Filter $ouFilter -SearchScope Subtree -ErrorAction Stop
    if ($ous) { $targets += $ous }
} catch {
    Write-Warning "Error while searching OU nodes: $_"
}

# 2) Optionally find CN containers named exactly as TargetNodeName
if ($IncludeCN) {
    try {
        # Using LDAP filter to locate containers by name; includes CN= entries
        $cns = Get-ADObject @adParams `
               -LDAPFilter "(name=$TargetNodeName)" `
               -SearchScope Subtree `
               -Properties objectClass,distinguishedName `
               -ErrorAction Stop
        if ($cns) { $targets += $cns }
    } catch {
        Write-Warning "Error while searching CN nodes: $_"
    }
}

# Remove duplicates by DN (some environments might return overlapping objects)
$targets = $targets | Sort-Object DistinguishedName -Unique

if ($targets.Count -eq 0) {
    Write-Warning "No AD nodes named '$TargetNodeName' were found."
    return
}

Write-Host ("Found {0} target node(s) named '{1}'" -f $targets.Count, $TargetNodeName) -ForegroundColor Green

# ==========================
# COUNT COMPUTERS PER TARGET
# ==========================
$result = @()

foreach ($t in $targets) {
    try {
        # Distinguished Name of the target node (search base)
        $dn  = [string]$t.DistinguishedName

        # Parse location as the element directly above the target node in DN
        $loc = Get-LocationFromDN -dn $dn -nodeName $TargetNodeName

        # Force array with @() so that .Count returns an integer reliably
        # Filter only computer objects
        $computers = @( Get-ADComputer @adParams -SearchBase $dn -SearchScope Subtree -LDAPFilter '(objectClass=computer)' -ErrorAction Stop )

        # Int conversion (should already be int, but keep defensive)
        $countInt = [int]$computers.Count

        # Accumulate results
        $result += [pscustomobject]@{
            DN       = $dn
            Location = $loc
            Count    = $countInt
        }
    } catch {
        Write-Warning "Error processing '$($t.DistinguishedName)': $_"
    }
}

if ($result.Count -eq 0) {
    Write-Warning "No computers found under nodes named '$TargetNodeName'."
    return
}

# ==========================
# TOTALS & PER-LOCATION SUMMARY
# ==========================
# Unpack Count as integers and sum defensively
$total = ($result | Select-Object -ExpandProperty Count | Measure-Object -Sum).Sum

# Group by Location and sum counts per group
$perLocation =
    $result |
    Group-Object Location |
    ForEach-Object {
        [pscustomobject]@{
            Location = $_.Name
            Total    = ($_.Group | Select-Object -ExpandProperty Count | Measure-Object -Sum).Sum
        }
    } |
    Sort-Object Location

# ==========================
# CONSOLE OUTPUT
# ==========================
Write-Host ""
Write-Host ("TOTAL computers under nodes '{0}': {1}" -f $TargetNodeName, $total) -ForegroundColor Magenta
Write-Host ("Summary by location (above '{0}'):" -f $TargetNodeName) -ForegroundColor Yellow
$perLocation | Format-Table -AutoSize

# ==========================
# TXT REPORT
# ==========================
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# Build report lines
$lines = @(
    "AD report for target node name: '$TargetNodeName'",
    "Generated: $timestamp",
    "",
    "TOTAL computers under '$TargetNodeName': $total",
    "",
    "Summary by location (location = DN element directly above '$TargetNodeName'):"
)

# Append each location line "Location : Total"
$lines += ($perLocation | ForEach-Object { "$($_.Location) : $($_.Total)" })

# Save report
$lines | Out-File -FilePath $TxtSummary -Encoding UTF8

Write-Host ""
Write-Host "Report saved: $TxtSummary" -ForegroundColor Green
