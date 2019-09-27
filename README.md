# Outlook Signature Autoconfig

Auto-generate and auto-configure Outlook signatures for your AD/LDAP users.

## How to use?

### Generate (`generate-signature.vbs`)

Run manually once for every relevant change in your AD. This will fetch all user info from AD, and generate signature file(s) (`.htm`) for each one of them, depending on which signature groups (`CN=assinaturas-ANYTHING`) they are associated with. The source directory (for templates) and target directory (for generated files) can be adjusted in `app_config.vbs`.

### Auto-configure (`outlook-signature-autoconfig.vbs`)

Run manually or via GPO. This will copy the user generated signature file(s) from the specified remote directory to a local directory created inside the user roaming directory (e.g.: `C:\Users\johndoe\Appdata\Roaming\TeltecSolutions\`). After that, it will update the default path for Outlook signature files, and set the default Outlook signature file to one of those that were copied. If it's already set to one of them, it will not change it.

## Features

- Decides which template to use based on the AD groups the user belongs to (signature groups);
- Allow a user to be part of multiple signature groups. In this case, the user will have multiple signature files, and the default signature will be kept intact if it's present in the remote signature folder, otherwise one of those will be chosen (normally the last one by alphabetical order).
- Handles users that belong to no signature groups;
- Validates template file existence;
- Log all errors to `stderr` when running on console (e.g.: via `cscript`, not `wscript`);

## Notes

- It's important to specify `\\host\path` instead of `X:\path` if you're running `outlook-signature-autoconfig.vbs` via GPO during user login. This is because `X:` is a network mapped unit that may be mapped via GPO, so it might not yet be mounted when our script runs;

## Screenshot (example template)

![Example Template](Templates/example/screenshot.png)
