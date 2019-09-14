# Powershell module for
# Computer associates Service Desk Manager
# By Xavier Plantefeve, 2019
# MIT License

#region enums & types

# I found no odfficial list in any documentation,
# so thevalues for this enum are taken from links
# found during my tests.
enum CasdObject {
    actbool
    chg
    cmth
    cnt
    cnt_role
    contact_handling
    cr
    cr_wf
    ctp
    dept
    grp
    grpmem
    iss
    iss_wf
    lr
    noturg
    org
    qp_chg_cnt
    qp_cr_cnt
    qp_iss_cnt
    wf
    wrkshft
    no_contract_sdsc
    zLocation
    lrel_att_cntlist_macro_ntf
    lrel_cenv_cntref
    lrel_dist_cntlist_mgs_ntf
    lrel_notify_list_cntchgntf
    lrel_notify_list_cntissntf
    lrel_notify_list_cntntf
    lrel_ntfr_cntlist_att_ntfrlist
    lrel_svc_grps_svc_chgcat
    lrel_svc_grps_svc_isscat
    lrel_svc_grps_svc_pcat
    lrel_svc_grps_svc_wftpl
    lrel_svc_locs_svc_groups
    z_lrel_ci_developers
    z_lrel_ci_key_users
    CI_ACTIONS_ALTERNATE
}
#endregion

#region utilities

# While most functions can be directed to use a
# specific API URL or Session Key, those two values
# are kept in script (so module here) scope variables
# and manipulated with the following four functions

function Set-BaseUrl {
    <#
    .SYNOPSIS
        Sets the base API url for subsequent use by functions.
    .INPUTS
        String
    .OUTPUTS
        None
    #>
    # The base API URL
    param(
        [parameter(Mandatory,Position=0,ValueFromPipeline)]
        [string]$BaseUrl
    )

    $Script:BaseUrl = $BaseUrl
}

function Set-AccessKey {
    <#
    .SYNOPSIS
        Sets the access key for subsequent use by functions.
    .INPUTS
        String
    .OUTPUTS
        None
    #>
    # The access key, see New-Session
    param(
        [parameter(Mandatory,Position=0,ValueFromPipeline)]
        [string]$AccessKey
    )

    $Script:AccessKey = $AccessKey
}

function Get-BaseUrl {
    <#
    .SYNOPSIS
        Gets the currently set base API url.
    .INPUTS
        None
    .OUTPUTS
        String
    #>
    $Script:BaseUrl
}

function Get-AccessKey {
    <#
    .SYNOPSIS
        Gets the currently set base API url.
    .INPUTS
        None
    .OUTPUTS
        String
    #>
    $Script:AccessKey
}

filter Out-Date {
    <#
    .SYNOPSIS
        Converts CASD Manager date to DateTime
    .DESCRIPTION
        Dates in CASD Manager are kept as number of seconds since
        the EPOCH, UTC time. The Out-Date filters converts
        CASD date to the DateTime type.
    .EXAMPLE
        PS C:\> 50105101 | Out6Dqte
        Show a legible version of the date.
    .INPUTS
        Int
    .OUTPUTS
        DateTime
    #>
    ([DateTime]'1970-01-01').Add(
        [DateTime]::Now - [datetime]::UtcNow
    ).AddSeconds($_)
}

# I have but one system to test, so the following will
# likely need revisions on other environments.
filter Out-CorrectEncoding {
    <#
    .SYNOPSIS
        Converts CASD Manager strings to the proper encoding
    .DESCRIPTION
        Deep down, the module uses PS REST cmdlets. In Windows 10    
        those cmdlets get the encoding wrong. The Out-CorrectEncoding 
        filter corrects that.
    .EXAMPLE
        PS C:\> 'La curiositÃ© est un vilain dÃ©faut... | Out-CorrectEncoding
        Show a legible version of the string.
    .INPUTS
        String
    .OUTPUTS
        String
    #>
    [text.encoding]::UTF8.GetString([text.encoding]::Default.GetBytes($_))
}

#endregion

#region main bricks
function Invoke-ApiRestMethod {
    <#
    .SYNOPSIS
        The Invoke-ApiRestMethod cmdlet is a wrapper around Invoke-RestApiMethod
    .DESCRIPTION
        The Invoke-ApiRestMethod cmdlet wraps the needed functionnality of
        Invoke-RestApiMethod while taking care of the API URL and access key.
    .INPUTS
        None
    .OUTPUTS
        Object
    .NOTES
        This cmdlet will be updated to handle the methods that
        will be needed down the line.
    #>
    [CmdletBinding(DefaultParameterSetName='UrlInfo')]
    Param(
        # The type of object to retrieve. This is defined in the CasdObject enum.
        [Parameter(Mandatory,position=0,ParameterSetName='UrlInfo')]
        [CasdObject]$Object,

        # The REST method to use. The list will be expanded as needed.
        [Parameter(position=1,ParameterSetName='UrlInfo')]
        [Parameter(position=0,ParameterSetName='URL')]
        [ValidateSet('Get','Post')]
        [string]$Method = 'Get',

        # The URL parameters, if needed.
        [Parameter(ParameterSetName='UrlInfo')]
        [string[]]$Parameters,

        # The URL for the rest method can be given as is.
        [Parameter(ParameterSetName='URI')]
        [string]$Uri,

        # The HTTP headers to send.
        [HashTable]$Headers,
        # The HTTP body for POST requests.
        [string]$Body,
        # The format of the body.
        [ValidateSet('application/json','application/xml')]
        [string]$Accept = 'application/json',
        
        # The base API URL. Defaults to the one set up with Set-BaseUrl
        [Parameter(ParameterSetName='UrlInfo')]
        [string]$BaseUrl = (Get-BaseUrl),
        # The credentials to use, for session creation. Any subsequent
        # request will use the then created access key
        [PSCredential]$Credential,
        # The access key, given at session creation
        [string]$AccessKey = (Get-AccessKey)
    )
    if ($null -eq $PSBoundParameters['Headers']) {
        $Headers = @{}
    }

    $Headers['Accept'] = $Accept

    # If the URL wasn't given through the -URI parameter,
    # we build it from the other ones.
    if ($PSBoundParameters['Object']) {
        $Object = $Object.ToLower()
        $Uri = $BaseUrl,$Object -join '/'
        if ($PSBoundParameters['Parameters']){
            $Uri = $Uri,($Parameters -join '&') -join '?'
        }
    }

    # $arguments will be splatted to Invoke-RestMethod
    $arguments = @{
        Uri = $Uri
        Method = $Method
    }

    if ($PSBoundParameters['Body']) {
        $arguments['Body'] = $Body
    }

    if ($AccessKey) {
        $Headers['X-AccessKey'] = $AccessKey
    }

    # Building the BASIC auth header,
    # for the creation of a new session.
    if ($PSBoundParameters['Credential']) {
        $AuthString = "{0}:{1}" -f (
            $Credential.Username,
            $Credential.GetNetworkCredential().Password
        )
        $AuthBytes  = [System.Text.Encoding]::Ascii.GetBytes($AuthString)
        $Headers['Authorization'] =
            'Basic {0}' -f [Convert]::ToBase64String($AuthBytes)
    }

    $arguments['Headers'] = $Headers
    
    Invoke-RestMethod @arguments
}

function New-Session {
    <#
    .SYNOPSIS
        Creates a new session
    .DESCRIPTION
        The New-Session cmdlet opens a new API session, outpouts
        the session information, and sets the session key for
        subsequent API use.
    .EXAMPLE
        PS C:\> New=Session -Credential (Get-Credential) -BaseUrl 'http://casd.contoso.com:8050/caisd-rest'
        Creates a new session
    .INPUTS
        None
    .OUTPUTS
        Object
    .NOTES
        The output object has three properties:
        - The access key
        - The end date and time of the session
        - The url of the session object
    #>
    param(
        [Parameter(Position=0)]
        [String]$BaseUrl,
        [Parameter(Mandatory)]
        [PSCredential]$Credential
    )

    if ($PSBoundParameters['BaseUrl']) {
        Set-BaseUrl -BaseUrl $BaseUrl
    }

    $arguments = @{
        Object     = 'rest_access'
        Method     = 'Post'
        Body       = ConvertTo-Json @{'rest_access' = $null}
        Accept     = 'application/json'
        Credential = $Credential
        Headers    = @{'Content-Type' = 'application/json'}

    }

    $Session = Invoke-ApiRestMethod @arguments | select -ExpandProperty rest_access
    
    Set-AccessKey -AccessKey $Session.access_key

    $properties = @(
        @{l='AccessKey';e={$_.access_key}}
        @{l='ExpirationDate';e={$_.expiration_date | Out-Date}}
        @{l='Link';e={$_.link}}
    )
    $Session = $Session | select $properties
    $Session.Link |
        Add-Member -MemberType ScriptMethod -Name ToString -Value {$this.'@href'} -Force
    
    $Session
}


function Get-ApiObject {
    param(
        [parameter(ParameterSetName='UriInfo',Mandatory,Position=0)]
        # [ValidateSet('cnt','grp','rest_access')]
        [CasdObject]$ObjectType,
        [Parameter(ParameterSetName='UriInfo')]
        [string]$Id,
        [Parameter(ParameterSetName='Uri',Mandatory)]
        [string]$Uri,
        [string[]]$Attributes = '*',
        [Parameter(ParameterSetName='UriInfo')]
        [Alias('WC')]
        [string]$WhereClause,
        [Parameter(ParameterSetName='UriInfo')]
        [Int]$Page,
        [Parameter(ParameterSetName='UriInfo')]
        [Int]$PageSize = 25,
        [ValidateNotNullOrEmpty()]
        [string]$BaseUrl = (Get-BaseUrl),
        [string]$AccessKey = (Get-AccessKey),
        [switch]$raw
    )

    $Headers = @{}

    if ($PSBoundParameters['Uri']) {
        $arguments = @{
            Uri = $Uri
        }
    } else {
        if ($PSBoundParameters['Id']) {
            $Uri = $BaseUrl,$ObjectType,$Id -join '/'
        } else {
            if ($PSBoundParameters['WhereClause']) {
                [Array]$UriParameters += ( 'WC={0}' -f [System.Uri]::EscapeDataString($WhereClause) )
                Write-Verbose ($UriParameters | Out-String)
            }
            if ($PSBoundParameters['Page']) {
                [Array]$UriParameters += ( 'start={0}' -f ( 1 + ( ( $Page -1 ) * $PageSize) ) )
                [Array]$UriParameters += ( 'size={0}' -f $PageSize )
            }

            $Uri = ( $BaseUrl,$ObjectType -join '/' ),( $UriParameters -join '&') -join '?'
            Write-Verbose "URI: $Uri"
        }
    }

    if ($PSBoundParameters['AccessKey']) {
        [Hashtable]$Headers += @{'X-AccessKey' = $AccessKey}
    }

    Write-Verbose $Headers | Out-String
    [Hashtable]$Headers += @{'X-Obj-Attrs' = $Attributes -join ','}

    $arguments = @{
        Uri = $Uri
        Headers = $Headers
    }

    $response = Invoke-ApiRestMethod @arguments
    
    if ( $collection = $response.psobject.Properties | ? name -Match '^collection_' ) {
        if ( $collection.Value.'@COUNT' -gt 0) {
            $collection.Value | select -exp ($collection.Name -replace '^collection_')
            if ( -not $PSBoundParameters['Page'] ) {
                while ($next = $collection.Value.link | ? '@rel' -eq 'next') {
                    Write-Verbose $next.'@href'
                    $response = Invoke-ApiRestMethod -Uri $next.'@href' -Headers $Headers
                    $collection = $response.psobject.Properties | ? name -Match '^collection_'
                    $collection.Value | select -exp ($collection.Name -replace '^collection_')
                }
            }
        }
    } else {
        $response | select -exp ( $response.psobject.properties.name )
    }
}

filter Out-UserObject {
    $out = New-Object -TypeName psobject -Property @{
        Id = $_.'@id'
        Name = $_.'@COMMON_NAME'
        class = $_.producer_id
    }

    Write-Verbose ('found {0}' -f $_.producer_id)
    if ( $properties = Get-Variable -Scope Script -Name ( 'properties_{0}' -f $_.producer_id ) -ValueOnly -ea 0 ) {
        Write-Verbose ('{0} has properties' -f $_.producer_id)
        foreach ($prop in $properties) {
            Write-Verbose ('properties: {0}' -f $prop['Name'])
            if ( -not $prop.ContainsKey('MemberType') ) { $prop['MemberType'] = 'NoteProperty' }
            $out | Add-Member -MemberType $prop['MemberType'] -Name $prop['Name'] -Value ( & $prop['value'] $_ )
        }
        Write-Verbose ('properties: {0}' -f ( $out.psobject.Properties.name | Out-String ) )
    }

    $out.psobject | Add-Member -MemberType NoteProperty -Name ApiObject -Value $_
    $out.psobject | Add-Member -MemberType NoteProperty -Name CachedProperties -Value @{}
    $out.psobject.TypeNames.Insert(0,( 'casd_{0}' -f $_.producer_id ))

    if ( $DefaultDisplaySet = Get-Variable -Scope Script -Name ( 'DefaultDisplaySet_{0}' -f $_.producer_id ) -ValueOnly -ea 0 ) {
        Write-Verbose ('{0} has displayset' -f $_.producer_id)
        $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet(‘DefaultDisplayPropertySet’,[string[]]$defaultDisplaySet)
        $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
        $out | Add-Member MemberSet PSStandardMembers $PSStandardMembers
    }

    Write-Verbose ('properties: {0}' -f ( $out.psobject.Properties.name | Out-String ) )

    $out
}

#endregion

#region Display sets & properties

$script:DefaultDisplaySet_cnt = 'Name','UserId','email','Id'

$script:properties_cnt = @(
    @{name='userid';value = { param($o) ; $o.userid }}
    @{name='email';value = { param($o) ; $o.email_address }}
    @{name='phone';value = { param($o) ; $o.phone_number -split ' , ' }}
    @{
        membertype = 'ScriptProperty'
        name       = 'Cccchanges'
        value      = { {
            if ( -not $this.psobject.CachedProperties.ContainsKey('changes') ) {
                $this.psobject.CachedProperties['changes'] = ( Get-ApiObject -Uri $this.psobject.ApiObject.all_chg.link.'@href' | Out-UserObject )
            }
            $this.psobject.CachedProperties['changes']
        } }
    }
)

$script:properties_chg = @(
    @{name='Summary';value = { param($o) ; $o.summary | Out-CorrectEncoding }}
    @{name='Opened';value = { param($o) ; $o.open_date | Out-Date }}
)

$script:DefaultDisplaySet_cr = 'Request', 'Status', 'EndUser',
'Summary', 'Created', 'Modified', 'Assignee', 'group', 'priority',
'category', 'Service', 'Building'

$script:properties_cr = @(
    @{name = 'Request'  ; value = { param($o) ; $o.ref_num } }
    @{name = 'Status'   ; value = { param($o) ; $o.'@COMMON_NAME' } }
    # 'VIP'
    @{name = 'EndUser'  ; value = { param($o) ; $o.customer.'@COMMON_NAME' } }
    @{name = 'Summary'  ; value = { param($o) ; if (($o.summary | Out-CorrectEncoding).length -lt 55) { $o.summary | Out-CorrectEncoding } else { '{0} ...' -f ($o.summary | Out-CorrectEncoding).Substring(0, 55) } } }
    @{name = 'Created'  ; value = { param($o) ; $o.open_date | Out-Date } }
    @{name = 'Modified' ; value = { param($o) ; $o.last_mod_dt | Out-Date } }
    @{name = 'Assignee' ; value = { param($o) ; $o.assignee.'@COMMON_NAME' } }
    @{name = 'Group'    ; value = { param($o) ; $o.group.'@COMMON_NAME' } }
    @{name = 'Priority' ; value = { param($o) ; $o.priority.'@COMMON_NAME' } }
    @{name = 'Category' ; value = { param($o) ; $o.category.'@COMMON_NAME' } }
    @{name = 'Service'  ; value = { param($o) ; Get-ApiObject -Uri $o.customer.link.'@href' -ea 0 | select -exp dept -ea 0| select -exp '@COMMON_NAME' -ea 0 } }
    @{name = 'Building' ; value = { param($o) ; $o.zBuilding.'@COMMON_NAME' } }
)
#endregion
