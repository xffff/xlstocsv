' Convert all XLS files in a folder to CSV

Dim oExcel
Dim oBook
Dim folderPath

folderPath	= Wscript.Arguments.Item(0)
Set oFSO	= CreateObject("Scripting.FileSystemObject")
Set oFolder	= oFSO.GetFolder(folderPath)
Set cFiles	= oFolder.Files
Set oShell	= CreateObject("Wscript.Shell")
Set oExcel	= CreateObject("Excel.Application")

For Each oFile in cFiles
   If UCase(oFSO.GetExtensionName(oFile.Name)) = "XLS" Then      
      Set oBook = oExcel.Workbooks.Open(folderPath & oFile.Name)
      oBook.SaveAs folderPath & oFile.Name & ".csv", 6
      oBook.Close False
   End If
Next

oExcel.Quit