#Requires -PSEdition Core -Version 7
using namespace System.Globalization
using namespace Microsoft.PowerShell.Commands

[CmdletBinding(DefaultParameterSetName = "ApiWebRequest")]
param(
   [Parameter(ValueFromPipeline, Position = 0, ParameterSetName = "ApiWebRequest")]
   [BasicHtmlWebResponseObject] $Response,
   [Parameter(ParameterSetName = "Debugging")]
   [PsCustomObject[]] $TransactionsFromApi,
   [IO.FileInfo] $OutputFile = (New-TemporaryFile),
   [Switch] $SaveApiResponse
)

$ErrorActionPreference = 'Stop'
$apiCulture = [CultureInfo]::new("en-US")
$targetCulture = [CultureInfo]::new("nl-NL")

$TransactionsFromApi ??= $Response.Content | ConvertFrom-Json

$transactionEntries = $TransactionsFromApi `
    | Select-Object -Property *, `
        @{ Name="ParsedSourceAmount"; Expression={[double]::Parse($_.sourceAmount, $apiCulture) } },
        @{ Name="ParsedBillingAmount"; Expression={[double]::Parse($_.billingAmount, $apiCulture) } } `
    | Select-Object -Property `
        @{ Name="TransactionDate"; Expression={$_.transactionDate } },
        @{ Name="Description"; Expression={($_.countryCode.Trim()) `
            ? ((($_.description.Split(' ')) | Select-Object -SkipLast 2) -Join " ").ToUpperInvariant() `
            : $_.description.ToUpperInvariant() } },
        @{ Name="MerchantCatagory"; Expression={$_.merchantCategoryCodeDescription } },
        @{ Name="CountryCode"; Expression={$_.countryCode.Trim() } },
        @{ Name="CardAcceptorIdentity"; Expression={($_.countryCode.Trim()) `
            ? (($_.Description.Split(' ')) | Select-Object -SkipLast 1 | Select-Object -Last 1 ) `
            : $null } },
        @{ Name="SourceAmount"; Expression={"$($_.ParsedSourceAmount.ToString("F", $targetCulture) ) $($_.sourceCurrency)" } },
        @{ Name="Deposit_BillingAmount"; Expression={($_.ParsedBillingAmount -lt 0 ) `
            ? ([Math]::Abs($_.ParsedBillingAmount)).ToString("F", $targetCulture) `
            : $null } }, 
        @{ Name="Withdrawal_BillingAmount"; Expression={($_.ParsedBillingAmount -gt 0 ) `
            ? $_.ParsedBillingAmount.ToString("F", $targetCulture) `
            : $null } }, 
        @{ Name="BillingCurrency"; Expression={$_.billingCurrency } },
        @{ Name="ExchangeRateStr"; Expression={($_.sourceCurrency -ne $_.billingCurrency) `
            ? "Wisselkoers $($_.sourceCurrency)" `
            : $null } },
        @{ Name="ExchangeRate"; Expression={($_.sourceCurrency -ne $_.billingCurrency) `
            ? ([Math]::Round(($_.ParsedSourceAmount / $_.ParsedBillingAmount), 5)) `
            : $null } }

$transactionEntries `
    | Sort-Object @{ Expression={Get-Date -Date $_.transactionDate }; Ascending=$true } `
    | ConvertTo-Csv -Delimiter ";" -NoTypeInformation `
    | Out-File $OutputFile.FullName -Force -Verbose:$true

if ($SaveApiResponse.IsPresent) {

    $TransactionsFromApi `
        | ConvertTo-Csv -Delimiter ";" -NoTypeInformation `
        | Out-File (Join-Path $OutputFile.Directory.FullName "$($OutputFile.BaseName)_FromApi$($OutputFile.Extension)") -Force -Verbose:$true
    
    $TransactionsFromApi `
        | ConvertTo-Json -Depth 100 `
        | Out-File (Join-Path $OutputFile.Directory.FullName "$($OutputFile.BaseName)_FromApi.json") -Force -Verbose:$true
}