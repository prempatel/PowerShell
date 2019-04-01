param (
  [Parameter(Mandatory=$true,position=0)]
  $MountDir = "E:\temp\Mount",

  [Parameter(Mandatory=$true,position=1)]
  $FirstWim = "C:\Wim Backup\Win10_OfficeAdobe.wim",

  [Parameter(Mandatory=$true,position=2)]
  $SecondWim = "E:\MDTRef\Captures\REF-WIN10l\REF-WIN10-Office.wim"
)

$Name = "Windows10WithOffice"

#Mount first wim
& dism /Mount-wim /Wimfile:"C:\WimLab\OfficeAdobe\install.wim" /index:1 /mountdir:"C:\Mount"

#Append Image
& dism /append-image /imagefile:"C:\WimLab\OfficeAdobe\install.wim" /CaptureDir:"C:\Mount" /Name:"GP15 Book"

#Cleanup 
& dism /unmount-image /mountdir:"C:\Mount" /discard

#Show Wim Info
& dism /get-imageinfo /imagefile:"C:\WimLab\OfficeAdobe\install.wim"

imagex /info "C:\WimLab\OfficeAdobe\install.wim" 6 "Windows 10 - GP15 Office16 Adobe" "Windows 10 with GP15, Office 16, and Adobe Reader"


imagex /info "C:\WimLab\OfficeAdobe\install.wim" 6 "Windows 10 - GP15 Office16 Adobe" "Windows 10 with Office 16 and Adobe Reader"

#Split first wim
#Add to USB and test