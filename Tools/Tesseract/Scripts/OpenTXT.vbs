Dim WshShell
Dim fso
Dim ShellCommand, ExePath, ExeCommand, Outpath, OutName, PDFName, strRetVal, PDF2TXT
Dim re
Dim xl

PDFName = Wscript.Arguments(0)
Set fso = CreateObject("Scripting.FileSystemObject")
OutName = fso.GetBaseName(PDFName)
ExePath = "cmd.exe /C C:\Macros\Tools\PDF2TXT\bin\pdf2txt.exe "
Outpath = fso.GetParentFolderName(PDFName)
PDF2TXT = Outpath & "\" & OutName & ".txt"
ExeCommand = "-table -fixed 3 -nopgbrk " & CMDPath(PDFName) & " " & CMDPath(PDF2TXT)
ShellCommand = ExePath & ExeCommand
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run ShellCommand, 0, True

Dim oFile: Set oFile = fso.GetFile(PDF2TXT)
With oFile.OpenAsTextStream()
        strRetVal = .Read(oFile.Size)
        .Close
End With


Set re = New RegExp
    With re
        .Pattern = "^(?:[\t ]*(?:\r?\n|\r))+"
        .IgnoreCase = False
        .MultiLine = True
        .Global = True
        With oFile.OpenAsTextStream(2)
                .Write(re.Replace(strRetVal, ""))
                .Close
        End With
	End With
	
Set xl = CreateObject("Excel.application")

xl.Application.Workbooks.Open PDF2TXT
xl.Application.Visible = True


Function CMDPath(FilePath)

If InStr(FilePath, " ") <> 0 Then
    CMDPath = """" & FilePath & """"
Else
    CMDPath = FilePath
End If

End Function