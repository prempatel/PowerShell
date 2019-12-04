#Reopen application if closed

Invoke-Item "C:\Users\ppatel\Work Folders\Desktop\TAKT Tracking.accde"

while ($true) {
    $access = Get-Process msaccess
    $access.WaitForExit();
    Invoke-Item "C:\Users\ppatel\Work Folders\Desktop\TAKT Tracking.accde"
}
