. ./psCASD-private.ps1

if (get-module pCASD) { Remove-Module psCASD } ; Import-Module .\psCASD.psm1 -Prefix $Prefix

if ( $null -eq (Get-EpsdAccessKey) ) {
    $Session = New-EpsdSession -BaseUrl $BaseUrl -Credential $cred
    $Session
}

if ( $null -eq $TST ) {
    $TST = Get-EpsdApiObject -Uri $TstUrl
}


$t1 = $tst | New-UserObject
$t1

function Get-Contact {
    param(
        [string[]]$Attributes = '*',
        [Alias('WC')]
        [string]$WhereClause,
        [Int]$Page,
        [Int]$PageSize = 25,
        [ValidateNotNullOrEmpty()]
        [string]$BaseUrl = (Get-EpsdBaseUrl),
        [string]$AccessKey = (Get-EpsdAccessKey)
    )

    Get-EPSDApiObject -ObjectType cnt -WhereClause $WhereClause -Attributes $Attributes | Out-UserObject
}


$wwks = Get-EpsdApiObject grp -WhereClause "last_name LIKE '$MyGroup%'" | select -ExpandProperty '@Id'

Get-EpsdApiObject cr -WhereClause "group = $wwks and assignee is NULL" | Out-UserObject | tee -Variable test | ft
