# Add PST to User Mailbox

Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null
$outlook = new-object -comobject outlook.application
$namespace = $outlook.GetNameSpace("MAPI")
dir "<unc path>" | % { $namespace.AddStore($_.FullName) }