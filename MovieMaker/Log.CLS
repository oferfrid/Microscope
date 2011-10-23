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

' MovieMaker is an ImagePro sofware for time lapsed serial microscope monitoring.
' Copyright 2011 Sivan Pearl published under GPLv3,
' this software was developed in Prof. Nathalie Q. Balaban's lab, at the Hebrew University , Jerusalem , Israel .
'
' This file is part of MovieMaker.
'
' MovieMaker is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' MovieMaker is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with MovieMaker. If not, see <http:'www.gnu.org/licenses/>.
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
    	fname = "MicroscopeLog" & Format(CDate(Now()), "YYYY-MM-DD-hh-nn-ss") & ".txt"
    End If
    fsCreateFileName=fname
End Function



Private Sub Class_Terminate()
    If mbFileHasBeenAssigned Then
        Close #miFileNumber
        mbFileHasBeenAssigned = False
    End If
    
End Sub
