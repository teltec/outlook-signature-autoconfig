Option Explicit

' NOTE: It's important to specify `\\host\path` instead of `X:\path` if you're running `outlook-signature-autoconfig.vbs` via GPO during user login. This is because `X:` is a network mapped unit that may be mapped via GPO, so it might not yet be mounted when our script runs;

' Directory where signature files are or will be stored
Const gConfigSignaturesSourceLocation = "\\file-server\Temp\TestOutlookSignatures"

' Directory where templates are stored
Const gConfigTemplatesSourceLocation = "\\file-server\Temp\TestOutlookSignatures\Templates"

' AD/LDAP details (server, base DN, user filter, user attributes to fetch)
Const gConfigLdapServer = "my.example.com"
Const gConfigLdapBaseDN = "OU=My Organiztion Unit,DC=example,DC=com"
Const gConfigLdapFilter = "(&(mail=*)(objectClass=organizationalPerson))"
Const gConfigLdapAttributes = "mail,memberOf,sAMAccountName,displayName,title,company,homePhone,telephoneNumber,mobile"
