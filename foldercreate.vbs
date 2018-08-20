Dim YearFolder, ReportingMonth
Set Args = WScript.Arguments

' For debugging - will echo count of command line parameters passed in and then each parameter in order
' WScript.Echo WScript.Arguments.Count
' For Each strArg in objArgs
'     WScript.Echo strArg
' Next

YearFolder=Args(0)
ReportingMonth=Args(1)

msgbox "YearFolder set as: " _
       & YearFolder & " and ReportingMonth set as: " & ReportingMonth

' Set oFSO = CreateObject("Scripting.FileSystemObject")

' If Not oFSO.FolderExists( "C:\data\Financial\Management Accounts\Month End\2018\1808\Reports\Aged Analysis\AP") Then
' Set objFolder = oFSO.CreateFolder("C:\data\Financial\Management Accounts\Month End\2018\1808\Reports\Aged Analysis\AP")
' End If

' If Not oFSO.FolderExists( "C:\data\Financial\Management Accounts\Month End\2018\1808\Reports\Aged Analysis\AR") Then
' Set objFolder = oFSO.CreateFolder("C:\data\Financial\Management Accounts\Month End\2018\1808\Reports\Aged Analysis\AR")
' End If


' Target file path - 
' R:\Financial\Management Accounts\Month End\2018\****\Reports
' R:\Financial\Management Accounts\Month End\2018\****\Reports\Aged Analysis\AP
' R:\Financial\Management Accounts\Month End\2018\****\Reports\Aged Analysis\AR


' Set oWS = WScript.CreateObject("WScript.Shell")

' Dim objFSO

' Set objFSO = CreateObject("Scripting.FileSystemObject")
' If NOT objFSO.FolderExists("C:\data\") Then
'     splitString = Split(userProfile, "\")
'     MsgBox("Local folder doesn't exist, creating...")
'     MsgBox("D:\" + splitString(2) + "\AppData\Roaming\Local")
'     WSHShell.Run "mkdir c:\FSO"
' End If
