InputSourcePath = "\\hqolymfl01\groupi$\631020\QuarterlyIncentivePayments\Input\"
InputDestPath = "\\hqolymfl01\groupi$\631020\QuarterlyIncentivePayments\Archive\Input\"

OutPutSourcePath = "\\hqolymfl01\groupi$\631020\QuarterlyIncentivePayments\Output\Incentives\"
OutPutDestPath = "\\hqolymfl01\groupi$\631020\QuarterlyIncentivePayments\Archive\Output\"

Set FSO = CreateObject("Scripting.FileSystemObject")
Set FLD = FSO.GetFolder(OutPutSourcePath)

dToday = Date
sToday = Right("0" & Day(dToday), 2) & MonthName(Month(dToday), True) & Year(dToday)

For Each fil In FLD.Files
    strNewName = OutPutDestPath & "Archived " & sToday & " " & fil.Name
FSO.MoveFile fil , strNewName 

Next

Set FLD = FSO.GetFolder(InputSourcePath)

For Each fil In FLD.Files
    strNewName = FSO.BuildPath(InputDestPath, "Archived " & sToday & " " & fil.Name)
    fil.Move strNewName
Next

msgbox "Files archived"