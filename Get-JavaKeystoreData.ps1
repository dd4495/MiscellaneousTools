<#
.SYNOPSIS
  Gets a formatted list of the contents of a java keystore
.DESCRIPTION
  Put this script in your java \bin folder to run it.
  Optionally outputs the keystore info to an xlsx file.
.EXAMPLE
  .\Get-JavaKeystoreData.ps1 -truststore F:\temp\ldap-server.truststore
  This will use the default password, and will write the results to the console. 
.EXAMPLE
  .\Get-JavaKeystoreData.ps1 -truststore F:\temp\ldap-server.truststore -xlsx -outFile F:\temp\keystore.xlsx
  This will use the default password and output an excel file
.PARAMETER truststore
  The path to the Java truststore file
.PARAMETER storepass
  The store password. Defaults to changeit
.PARAMETER xlsx
  If enabled, this switch will save the results to an excel file
.PARAMETER outFile
  The location of the outputted file. Only necessary if xlsx is enabled.
.NOTES
  Version:        1.0
  Author:         dd4495 
  Creation Date:  October 2016
#>
param (
    [Parameter(Mandatory,ValueFromPipeline,ValueFromPipelineByPropertyName,Position=0)]
    [string]$truststore,
    [Parameter(ValueFromPipeline,ValueFromPipelineByPropertyName,Position=1)]
    [string]$storepass = 'changeit',
    [Parameter(ValueFromPipeline,ValueFromPipelineByPropertyName,Position=2)]
    [switch]$xlsx,
    [Parameter(ValueFromPipeline,ValueFromPipelineByPropertyName,Position=3)]
    [string]$outFile = 'F:\temp\keystoreData.xlsx'
    )


$Certs = @()

$results = & .\keytool.exe -keystore $truststore -list -storepass $storepass
$tResults = $results[6..($results.Length -5)]
$fingerprints = for($i=1;$i -lt $tResults.Count;$i+=2){$tResults[$i]}
$names = for($i=0;$i -lt $tResults.Count;$i+=2){$tResults[$i]}

for ($i=0;$i -lt $names.Count;$i++) {
    $certName = $names[$i] -split (',',2) | Select-Object -Index 0
    $creationDate = (($names[$i] -split (',',4) | Select-Object -Index 1,2) -join ",").trim()
    $type = (($names[$i] -split (',',4) | Select-Object -Index 3) -replace ',','').trim()
    $encryption = ((($fingerprints[$i] -split (' ',4) | Select-Object -Index 2) -replace ':','') -replace '\(','') -replace '\)',''
    $prints = ($fingerprints[$i] -split (':',2) | Select-Object -Index 1).trim()

    $props = [ordered]@{'AliasName'       = $certName;
                        'CreationDate'    = $creationDate;
                        'Encryption'      = $encryption;
                        'Fingerprint'     = $prints;
                        'EntryType'       = $type;

    }
    $CertInfo = New-Object -TypeName psObject -Property $props
    $Certs += $CertInfo
}

if ($xlsx) {Write-Output $Certs | Export-Excel -Path $outFile}
else {$Certs}
