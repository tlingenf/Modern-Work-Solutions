
Import-Module Microsoft.PowerApps.PowerShell
Import-Module Microsoft.PowerApps.Administration.PowerShell

function Create-DlpPolicy {
  [CmdletBinding()]
  param(
    [parameter(ValueFromPipeline)]
    $environment,

    [parameter(Mandatory=$true)]
    $policyName
  )

  process {
    $policy = $null
    # Create a Power Platform DLP Policy
    try {
      $policy = Get-DlpPolicy -policyName $policyName
    }
    catch {
      $policy = New-AdminDlpPolicy -Name $policyName -Enabled $true
    }

    

  }
}

function Get-DlpPolicy {
  [CmdletBinding()]
  param(
    [parameter(Mandatory=$true)]
    $policyName
  )

  process {
    # Get a Power Platform DLP Policy
    $policy = Get-AdminDlpPolicy $policyName
    if ($policy) {
      $policy
    } else {
      throw "Policy $policyName not found"
    } 
  }
}

function Get-PolicyConfiguration {
  [CmdletBinding()]
  param (
    [parameter(Mandatory=$true)]
    [string]$ExcelPath,

    [parameter(Mandatory=$true)]
    [string]$PolicyName
  )
  
  begin {
    $EXCEL_COLUMN_
    $objExcel = New-Object -ComObject Excel.Application
    $objExcel.Visible = $false
    $objExcel.DisplayAlerts = $false
    $objWorkbook = $objExcel.Workbooks.Open($ExcelPath)
    $objWorksheet = $objWorkbook.Worksheets.Item($PolicyName)
  }
  
  process {
    $range = $objWorksheet.UsedRange
    for ([int]$index = 2; $index -le $range.Rows.Count; $index++) {
      $row = $range.Rows.Item($index)
      $row.Cells.Item(1).Text
    }
  }
  
  end {
    $objExcel.Quit()
  }
}