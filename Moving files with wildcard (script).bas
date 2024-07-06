Attribute VB_Name = "Module1"
Sub MoveWithWildcard()

Dim targetDir As String
Dim targetFolder As String
Dim FileName As String
Dim wsheet As Worksheet
Dim FSO As Object
Dim wshell As Object
Dim totalRows As Long

Set wsheet = ThisWorkbook.Sheets("Sheet1")
Set FSO = CreateObject("Scripting.Filesystemobject")
Set wshell = CreateObject("Wscript.shell")

'To count the value started from A5 to the bottom
totalRows = wsheet.Range("A5", wsheet.Range("A1").End(xlDown)).Rows.Count

'Return the current working dir
ActiveDir = wshell.CurrentDirectory

targetDir = wsheet.Range("B1") 'Range (B1) = C:\Users\xxxxx\Downloads
FolderName = wsheet.Range("B2") 'Range (B2) = Test
targetFolder = targetDir & "\" & FolderName & "\" 'C:\Users\xxxxx\Downloads\Test\

For i = 5 To ((5 + totalRows) - 1)

    FileName = "*" & wsheet.Range("A" & i) & "*"
    MovingFiles = ActiveDir & "\" & FileName
    
    FSO.MoveFile Source:=MovingFiles, Destination:=targetFolder

Next i

End Sub

