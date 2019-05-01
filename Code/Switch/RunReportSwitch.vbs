Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("\\hqolymfl01\groupi$\631020\QuarterlyIncentivePayments\Code\Switch\Switch.txt", 2)

objFile.Write "Run"

msgbox "Please be patient while we run your report.  You should recive an email in less than 10 minutes."