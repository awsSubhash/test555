Dim WshShell
Dim fso
Dim ShellCommand, ExePath, ExeCommand, Outpath, OutName, PDFName, MinutesElapsed

PDFName = Wscript.Arguments(0)    
Set fso = CreateObject("Scripting.FileSystemObject")
OutName = fso.GetBaseName(PDFName)
ExePath = "cmd.exe /C C:\Macros\Tools\PDF2TXT\bin\pdf2txt.exe "
Outpath = fso.GetParentFolderName(PDFName)
PDF2TXT = Outpath & "\" & OutName & ".txt"
ExeCommand = "-table -fixed 3 " & CMDPath(PDFName) & " " & CMDPath(PDF2TXT)
ShellCommand = ExePath & ExeCommand
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run ShellCommand, 0, True

Set objFileToRead = fso.OpenTextFile(PDF2TXT,1)
strFileText = objFileToRead.ReadAll()
pages = UBound(Split(strFileText, ""))
objFileToRead.Close
Set objFileToRead = Nothing

MinutesElapsed = FormatDateTime((Timer - Wscript.Arguments(1)) / 86400,3)


Set xl = CreateObject("Excel.application")
xl.Application.Workbooks.Open PDF2TXT
xl.Application.Visible = True
trans = xl.WorksheetFunction.CountA(xl.ActiveWorkbook.Worksheets(1).Range("A:A")) - 1
Call UpdatePDFLog(OutName, "ENCODED/SCAN", 1, pages, trans, MinutesElapsed, xl.username)
MsgBox "Please Copy the file to worksheet"
Function CMDPath(FilePath)

If InStr(FilePath, " ") <> 0 Then
    CMDPath = """" & FilePath & """"
Else
    CMDPath = FilePath
End If

End Function

Function UpdatePDFLog(filename, bs, files, pages, trans, proce, username)

Dim URL, res, Objhttp, strURL

URL = "https://docs.google.com/forms/u/0/d/e/1FAIpQLSd8ZDJ5ZgBqrK3avZyp1Y0Tfo-uJbWk2pz6SBQWe2mqFlBrbg/formResponse"
res = "entry.1657816269=" & filename & "&entry.1303594870=" & bs & "&entry.1259760741=" & "&entry.131490742=" & files & "&entry.2001117128=" & pages & "&entry.1153954912=" & trans & "&entry.598109455=" & proce & "&entry.162966060=" & username

Const ForWriting = 2

strURL= URL & "?" & res
Set objHTTP = CreateObject("MSXML2.XMLHTTP") 
Call objHTTP.Open("GET", strURL, FALSE) 
objHTTP.Send
End Function