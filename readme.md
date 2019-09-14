# Powershell module for CA ServiceDesk Manager

This is very much a work in progress and in the current state
very basic. For the time being, it handle the creation of
a session ( New-Session ) and interogates the API thanks to
Get-ApiObject.

There should be a module manifest, a version number, proper
documentation and comments-based help, and Pester tests.
There isn't for now.

## About the internal names

Rather than using names like Do-CasdSomething, I chose to use
Do-Something, and  the module is to be imported with a prefix,
the default one being Casd. It was done because it is likely
that the CA ServiceDesk Manager is rebranded when installed.
It is the case where I'm writing this module.

I'm still trying to find a nice nomenclature for functions names.
Some names will likely change.

## Example

Works is being done to create objects easy to manipulate, and
the basic bricks for that are working. An example of this would
be:

An API record is requested with Get-ApiObject. This object can
pass through the Out-UserObject filter. This filter recognized
the type of object, and dresses it thanks to pre-written hashtable
containing the wanted properties, and the default set of properties.

```Powershell
# 'cnt' is the object type for an user, or contact.
$contact = Get-ApiObject -ObjectType cnt -WhereClause "userid='xplantefeve'"

$contact | Out-UserObject
```
That last line uses the following properties:
```Powershell
$script:DefaultDisplaySet_cnt = 'Name','UserId','email','Id'

$script:properties_cnt = @(
    @{name='userid';value = { param($o) ; $o.userid }}
    @{name='email';value = { param($o) ; $o.email_address }}
    @{name='phone';value = { param($o) ; $o.phone_number -split ' , ' }}
    @{
        membertype = 'ScriptProperty'
        name       = 'Changes'
        value      = { {
            if ( -not $this.psobject.CachedProperties.ContainsKey('changes') ) {
                $this.psobject.CachedProperties['changes'] = ( Get-ApiObject -Uri $this.psobject.ApiObject.all_chg.link.'@href' | Out-UserObject )
            }
            $this.psobject.CachedProperties['changes']
        } }
    }
)
```
The object has the properties userid, email, phone and changes, but this
last one is not shown by default, it has to be expressely selected.

The user can then be displayed quickly. But if the changes is called, another
API call will be run to populate the properties. Additionaly, the result of
the call is cached and can be reused.
