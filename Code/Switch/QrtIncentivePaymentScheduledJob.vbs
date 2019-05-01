Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("G:\QuarterlyIncentivePayments\Code\Switch\Switch.txt", 1)
strContents = objFile.ReadAll
objFile.Close

If strContents = "Run" Then

	LogResult = "Run"
	Set oShell = CreateObject("WScript.Shell")
    	'run command'
    	Set oExec = oShell.Exec("C:\Program Files\R\R-3.4.4\bin\Rscript.exe G:\QuarterlyIncentivePayments\Code\Rcode.R")
	
	Set objFile = objFSO.OpenTextFile("G:\QuarterlyIncentivePayments\Code\Switch\Log.csv", 8)
	objFile.WriteLine LogResult & "," & Now & ","
	objFile.Close

	Set objFile = objFSO.OpenTextFile("G:\QuarterlyIncentivePayments\Code\Switch\Switch.txt", 2)
	objFile.Write "Dont Run"
	objFile.Close

Else
	LogResult  = "Not Run"
	Set objFile = objFSO.OpenTextFile("G:\QuarterlyIncentivePayments\Code\Switch\Log.csv", 8)
	objFile.WriteLine LogResult & "," & Now & ","
	objFile.Close
	
end if

