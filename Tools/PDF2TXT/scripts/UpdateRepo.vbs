Dim fso, BankName, OtherFor, PDFFiles, file_name, DestinationPath, SubFol1, SubFol2

BankName = Wscript.Arguments(0) 
OtherFor = Wscript.Arguments(1) 
PDFFiles = Wscript.Arguments(2)


DestinationPath = "W:\Shared With Me\PDF-Automation\Other Formats\"
SubFol1 = BankName
SubFol2 = OtherFor


file_name = Replace(Replace(Replace(Now, ":", ""), " ", ""),"-","")
Set fso = CreateObject("Scripting.FileSystemObject")
'Wscript.Echo DestinationPath
If Not fso.FolderExists(DestinationPath & SubFol1) Then fso.CreateFolder(DestinationPath & SubFol1)
DestinationPath = DestinationPath & SubFol1 & "\" & SubFol2 & "\"
If Not fso.FolderExists(DestinationPath) Then fso.CreateFolder(DestinationPath)
fso.CopyFile PDFFiles, DestinationPath & file_name & ".pdf"