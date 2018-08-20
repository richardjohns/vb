Set objShell = CreateObject("Wscript.Shell")

' Declare variables from parameters passed in from command line, and alert to confirm these.
Dim YearFolder, ReportingMonth, FULL_PATH
Set Args = WScript.Arguments
YearFolder=Args(0)
ReportingMonth=Args(1)
msgbox "YearFolder set as: " _
       & YearFolder & " and ReportingMonth set as: " & ReportingMonth

' Quit if not all parameters there, or parameters have letters, or are incorrect length.
If Wscript.Arguments.Count <> 2 Then
    Wscript.echo "Incorrect Parameters. Two required."
    Wscript.Quit
End If

' Create EOM folders from file path and variables.
FULL_PATH = "C:\data\Financial\Management Accounts\Month End\" + YearFolder + "\" + ReportingMonth + "\Reports\Aged Analysis\AP"
Set oFSO = CreateObject("Scripting.FileSystemObject")
BuildPath FULL_PATH
Sub BuildPath(ByVal Path)
If Not oFSO.FolderExists(Path) Then
BuildPath oFSO.GetParentFolderName(Path)
oFSO.CreateFolder Path
End If
End Sub

' File paths to create - 
' R:\Financial\Management Accounts\Month End\2018\****\Reports\Aged Analysis\AP
' R:\Financial\Management Accounts\Month End\2018\****\Reports\Aged Analysis\AR

' Command line version
' If Not oFSO.FolderExists("C:\data\Financial\Management Accounts\Month End\" + YearFolder + "\" + ReportingMonth + "\Reports\Aged Analysis\AP") Then
'     objShell.Run "cmd /c mkdir ""C:\data\Financial\Management Accounts\Month End\" + YearFolder + "\" + ReportingMonth + "\Reports\Aged Analysis\AP"""
' End If

' For debugging - will echo count of command line parameters passed in and then each parameter in order
' WScript.Echo WScript.Arguments.Count
' For Each strArg in objArgs
'     WScript.Echo strArg
' Next
