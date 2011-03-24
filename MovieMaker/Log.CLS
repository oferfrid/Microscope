VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
 END
Attribute VB_PredeclaredId = False
Attribute VB_Creatable = True
Attribute VB_Exposed = True
Attribute VB_GlobalNameSpace = False
Attribute VB_Name = "Log"
'Class name : Log
'Description : an easy an very usefull object to add loging to a batch application,saves you the hassle of writing this kind
'               of code over and over again.
'Author : Stefan De Prins
'Mail : stefan.de.prins@pandora.be

Option Explicit
Private miFileNumber As Integer
Private mbFileHasBeenAssigned As Boolean
Private mvarTitle As String
Private fname As String
Public Sub WriteLine(ByVal vsLine As String)
	If mbFileHasBeenAssigned Then
        Print #miFileNumber, vsLine
    Else
        MsgBox "No file has been assigned"
    End If
End Sub
Public Sub WriteLogEntery(ByVal mLine As String)
Dim vsLine As String
vsLine = Now()& mLine
    If mbFileHasBeenAssigned Then
        Print #miFileNumber, vsLine
    Else
        MsgBox "No file has been assigned"
    End If
End Sub

Public Sub WriteZLogEntery(ByVal zSendPos As Single)
Dim vsLine As String
'Dim z As Single
'ret=IpStageControl(GETZ,z)
vsLine = Now()&" Z is:" & vbTab & zSendPos
    If mbFileHasBeenAssigned Then
        Print #miFileNumber, vsLine
    Else
        MsgBox "No file has been assigned"
    End If
End Sub
Public Sub WriteLocLogEntery(ByVal xSendPos As Single,ByVal ySendPos As Single,ByVal zSendPos As Single)
Dim vsLine As String
vsLine = Now()&" X is:" & vbTab & xSendPos &" Y is:" & vbTab & ySendPos &" Z is:" & vbTab & zSendPos
    If mbFileHasBeenAssigned Then
        Print #miFileNumber, vsLine
    Else
        MsgBox "No file has been assigned"
    End If
End Sub
Public Sub CloseFile()
    If mbFileHasBeenAssigned Then
        Close #miFileNumber
        mbFileHasBeenAssigned = False
    Else
        MsgBox  "No file has been assigned"
    End If
End Sub

Public Sub WriteFooter()
    WriteLine "***********************************************"
    WriteLine "END " & mvarTitle & " ON " & Format(CDate(Now()), "YYYY-MM-DD hh:nn:ss")
    WriteLine "***********************************************"
End Sub

Public Sub WriteHeader()
    WriteLine "***********************************************"
    WriteLine "START " & mvarTitle & " ON " & Format(CDate(Now()), "YYYY-MM-DD hh:nn:ss")
    WriteLine "***********************************************"
End Sub

'opens a file with the given path for output, if the file allready exists it is open in append mode. 
Public Sub OpenFile(ByVal Path As String, ByVal vsTitle As String)
Dim lsFile As String

    Path = Trim(Path)
    miFileNumber = FreeFile
    lsFile = Path & fsCreateFileName
    If Dir(lsFile) = "" Then
        Open lsFile For Output As #miFileNumber
    Else
        Open lsFile For Append As #miFileNumber
    End If
    
    If Dir(lsFile) <> "" Then
        mbFileHasBeenAssigned = True
    Else
        MsgBox "File Creation failed"
    End If
    
    mvarTitle = vsTitle
End Sub

Public Sub CreateFile2(ByVal FileName As String)
    miFileNumber = FreeFile
    If Dir(FileName) = "" Then
        Open FileName For Output As #miFileNumber
    Else
        Open FileName For Append As #miFileNumber
    End If
End Sub
Private Function fsCreateFileName() As String
    If fname = "" Then
	fname = "MicroscopeLog" & Format(CDate(Now()), "YYYY-MM-DD") & ".txt"
End If
fsCreateFileName=fname
End Function
Private Sub Class_Terminate()
    If mbFileHasBeenAssigned Then
        Close #miFileNumber
        mbFileHasBeenAssigned = False
    End If
    
End Sub
