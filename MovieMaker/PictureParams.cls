VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PictureParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class name : PictureParams
'Description : a container for parameters involving lightning such as :phase/fluorescence, filter, exposure time etc.
'Author : Stefan De Prins
'Mail : stefan.de.prins@pandora.be
' needs to pay attention to the filters numbers- is it reflects the microscope settings
Option Explicit
Private exposureTime As Single 'im milliseconds
Private is_phase As Boolean
Private mFilter As Integer
Private framesFreq As Integer 'the number of frames this setting shoud be used -1 : every round, 2: every second round etc.
Private gain As Integer
Private focusOffset As Single 'to be higher or lower in the Z axis from focus found
Private empty_filter As Integer


Private Sub Class_Initialize()
focusOffset = 0
framesFreq = 1
empty_filter = 0
End Sub

Public Sub SetFrameFreq(freq As Integer)
    framesFreq = freq
End Sub

Public Function GetFrameFreq() As Integer
    GetFrameFreq = framesFreq
End Function
Public Sub SetExposureTime(sec As Single)
    exposureTime = sec
End Sub
Public Function GetExposureTime() As Single
    GetExposureTime = exposureTime
End Function

Public Sub SetPhase()
    is_phase = True
End Sub

Public Sub SetFluo()
    is_phase = False
End Sub
Public Function IsPhase() As Integer
    IsPhase = is_phase
End Function
Public Function IsFlour() As Integer
    IsFlour = False
    If is_phase = False Then
        IsFlour = True
    End If
End Function
Public Sub SetEmptyFilter()
    mFilter = empty_filter
End Sub
Public Function GetFilter() As Integer
    GetFilter = mFilter
End Function
Public Function SetFilter(ByVal filterNum As Single)
	mFilter =filterNum
End Function
Public Sub SetFocusOffset(ByVal offset As Single)
   focusOffset = offset
End Sub
Public Function GetFocusOffset() As Single
    GetFocusOffset = focusOffset
End Function
Public Sub SetGain(ByVal g As Integer)
    gain = g
End Sub
Public Function GetGain() As Integer
    GetGain = gain
End Function
'Return true is this picture should be taken in the given round.
Public Function shouldTakeFrame(ByVal roundNum As Integer) As Boolean
    Dim ans As Boolean
    Dim modulo As Integer
       modulo = roundNum Mod framesFreq
       shouldTakeFrame = False
    If modulo = 0 Then
        shouldTakeFrame = True
    End If
End Function
Public Sub copyPictureParams(src As PictureParams)
	If src.Is_Phase = True Then
		SetPhase
	Else
		SetFluo
	End If
	SetGain(src.GetGain)
	SetFilter(src.GetFilter)
	SetFocusOffset(src.GetFocusOffset)
	SetExposureTime(src.GetExposureTime)
	SetFrameFreq(src.GetFrameFreq)
End Sub
Public Function GetPicName() As String
	Dim pName As String
	If is_phase = True Then
		pName = "ph"
	Else
	pName = "filter_" & mFilter
	End If
	pName = pName$ &"_expo"&exposureTime
	If focusOffset<>0 Then
		pName = pName$ & "_offset"&focusOffset
	End If
	GetPicName = pName
End Function




