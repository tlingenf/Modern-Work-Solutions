param (
    [Parameter(Mandatory = $true)]
    [string]$GroupName,

    [Parameter(Mandatory = $false)]
    [string]$GroupAlias,

    [Parameter(Mandatory = $false)]
    [string]$PlanName
)

if (-not $GroupAlias) { $GroupAlias = "m365msgctr-notify"}
if (-not $PlanName) { $PlanName = "Message Center Tasks"}

Write-Output "Loading modules"
Import-Module Microsoft.Graph.Groups
Import-Module Microsoft.Graph.Files
Import-Module Microsoft.Graph.Planner

## only connect when not connected. Create a variable $connected = $true to bypass login
if (-not $connected) {
    Write-Output "Connecting to the Microsoft Graph ..."
    Connect-MgGraph -Scopes "https://graph.microsoft.com/Group.ReadWrite.All https://graph.microsoft.com/Files.ReadWrite.All https://graph.microsoft.com/Group.ReadWrite.All https://graph.microsoft.com/Tasks.ReadWrite"
    $connected = $true
}

$mgContext = Get-MgContext
$currentUser = Get-MgUser -UserId $mgContext.Account

$foundGroup = Get-MgGroup -Filter "DisplayName eq '$($GroupName)'"

if ($foundGroup) {
    Write-Output "$($GroupName) was found."
} else {
    Write-Output "$($GroupName) not found. Creating "
    $GroupOwner = $currentUser.Id
    $Owner = "https://graph.microsoft.com/v1.0/users/" + $GroupOwner
    $NewGroupParams = @{
        "displayName" = $GroupName
        "mailNickname"= $GroupAlias
        "description" = "Microsoft 365 Message Center & Roadmap"
        "owners@odata.bind" = @($Owner)
        "groupTypes" =  @(
                        "Unified"
                        )
        "mailEnabled" =  "true"
        "securityEnabled" = "true"
        "assignedLabels" = @()
    } 
    $foundGroup = New-MgGroup -BodyParameter $NewGroupParams
}

$foundPlan = Get-MgGroupPlannerPlan -GroupId $foundGroup.Id | Where-Object { $_.Title -eq $PlanName }

if ($foundPlan) {
    Write-Output "Plan found"
} else {
    Write-Output "Creating plan"
    New-MgPlannerPlan -Container $foundGroup.Id -Buckets ""

}