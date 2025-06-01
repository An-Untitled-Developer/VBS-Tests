' セッティング DIM/VAR, 関数宣言
' -------------------------------------------------------------
Set START = CreateObject("Wscript.Shell")

Dim PREDEF, OBJ, TXT
	PREDEF = "it's pronounced im-in-mi-no"
	Set OBJ = CreateObject("Scripting.FileSystemObject")
	TXT = START.ExpandEnvironmentStrings("%TEMP%") & "\imino.txt"

Set WRITE = OBJ.CreateTextFile(TXT, True)
	WRITE.Write PREDEF
WRITE.Close
' -------------------------------------------------------------
' GO

Dim GO
Set GO = START.Exec("notepad.exe """ & TXT & """")
Do While GO.Status = 0
WScript.Sleep 250
Loop

If OBJ.FileExists(TXT) Then
On Error Resume Next
OBJ.DeleteFile TXT
On Error GoTo 0
End If





