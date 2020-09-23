Attribute VB_Name = "Module1"
Sub WriteIni(File, Setting, NewInfo)
Dim NM As Label
Dim A_Line As String

On Error GoTo NM
Open "C:\OutPut.Dat" For Append As #1
Open File For Input As #2

Do While Not EOF(2)
Input #2, A_Line
If InStr(A_Line, Setting) Then
Print #1, Setting & NewInfo
Else
Print #1, A_Line
End If
Loop

Close #1
Close #2
Kill File
FileCopy "C:\OutPut.Dat", File
Kill "C:\OutPut.dat"
Exit Sub
NM:
MsgBox "Error writing to INI", vbSystemModal, "Error"
End Sub
