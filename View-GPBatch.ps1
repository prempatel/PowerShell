[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

$batchnumber = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Batch Number", "Release Batch")
while([string]::IsNullOrWhitespace($batchnumber)){
    $batchnumber = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Batch Number", "Release Batch")
}

Invoke-Sqlcmd -Query "select * from TBI..SY00500 WHERE BACHNUMB LIKE '$batchnumber%'" -ServerInstance "NJ-VSQL-01" | select BCHSOURC, BACHNUMB, MKDTOPST, BCHSTTUS

# Clear Batch status
#Invoke-Sqlcmd -Query "UPDATE TBI..SY00500 SET MKDTOPST = 0 , BCHSTTUS = 0 WHERE BACHNUMB LIKE '$batchnumber%'" -ServerInstance "NJ-VSQL-01"