VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PictureController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'#Uses "PictureParams.cls"
'#Uses "StageController.cls"
'Class name : PictureController
'Description : a class controlling the lighting
'Author : Stefan De Prins
'Mail : stefan.de.prins@pandora.be
' needs to pay attention to the filters numbers- is it reflects the microscope settings
Option Explicit
Private exposureTime As Single
Private gain As Integer
Private timeout As Integer
Private stgControl As New StageController
Private red_filter As Integer
Private cherry_filter As Integer
Private ko_filter As Integer
Private gfp_filter As Integer
Private yfp_filter As Integer
Private cfp_filter As Integer
Private focusOffset As Single
'Private lampVoltage As Integer

Private Sub Class_Initialize()
    gain = 255
    timeout = 20000
    red_filter = 0
    cherry_filter = 1
    ko_filter = 2
    gfp_filter = 3
    yfp_filter = 3
    cfp_filter = 0
    focusOffset = 0.0002
    'lampVoltage = -1

End Sub

Public Sub OpenPhaseShutter()
	'If (lampVoltage = -1) Then
	'	IpScopeGetPosition(SCP_LAMP1, lampVoltage)
	'End If
	'IpScopeSetPosition(SCP_LAMP1, lampVoltage)
	'Debug.Print "opening shuter light: " & lampVoltage
    ret = IpScopeControl(SCP_SETCURRSHUTTER, 0, 0, "", IPNULL)
    ret = IpScopeSetPosition(SCP_CURRSHUTTER, 1)
End Sub

Public Sub ClosePhaseShutter()
	'IpScopeGetPosition(SCP_LAMP1, lampVoltage)
	'IpScopeSetPosition(SCP_LAMP1, 0)
	'Debug.Print "closing shuter light: " & lampVoltage
    ret = IpScopeControl(SCP_SETCURRSHUTTER, 0, 0, "", IPNULL)
    ret = IpScopeSetPosition(SCP_CURRSHUTTER, 0)
    
End Sub
Public Sub OpenFluoShutter()
	ret = IpScopeControl(SCP_SETCURRSHUTTER, 0, 2, "", IPNULL)
    ret = IpScopeSetPosition(SCP_CURRSHUTTER, 1)
    End Sub
Public Sub CloseFluoShutter()
    ret = IpScopeControl(SCP_SETCURRSHUTTER, 0, 2, "", IPNULL)
    ret = IpScopeSetPosition(SCP_CURRSHUTTER, 0)
End Sub

Public Sub moveToFilterNum(ByVal filterNum As Integer)
    ret = IpScopeSetPosition(SCP_FWHEEL1, filterNum)
End Sub

Public Sub moveToFilter(ByVal picParams As PictureParams)
   Dim filterNum As Integer
   filterNum = picParams.GetFilter
   moveToFilterNum (filterNum)
End Sub

Public Sub SetExposureTime(ByVal picParams As PictureParams)
    Dim exposuresec As Single
    exposureTime = picParams.GetExposureTime()
    'Sets the exposure time in mseconds
    If exposureTime < 22 Then
        exposureTime = 22
    End If
    exposuresec = exposureTime / 1000
    IpDcamExposure (exposuresec)
   
End Sub

Public Sub SetGain(ByVal picParams As PictureParams)
    gain = picParams.GetGain()
    IpAcqControl(53, 1, gain)
    ret = IpDcamGain(gain)
End Sub
Public Sub PicAcquisition(ByVal picParams As PictureParams)

    If picParams.IsPhase Then
        PhaseAcquisition (picParams)
    Else
        FlouAcquisition (picParams)
    End If
End Sub

Public Sub PhaseAcquisition(ByVal picParams As PictureParams)
	Call moveToFilterNum(red_filter)
	Dim focusOffset As Single
    focusOffset = picParams.GetFocusOffset/1000
 	If focusOffset <> 0 Then
  	  Call stgControl.MoveZBy(focusOffset)
    End If
    '########
		Dim zz As Single
    	 ret = IpStageControl(GETZ, zz)
    	 Debug.Print("phaseAcquisition: taking picture at: "&zz)
    	 '########
    Call SetAcquisitionStart(picParams)
    Call triggerMode
    Call OpenPhaseShutter
    Call TriggerCamera
    Call ClosePhaseShutter
    Call TransferFrame
    Call SetAcquisitionEnd
    If focusOffset <> 0 Then ' return to focus
        stgControl.MoveZBy(-focusOffset)
    End If
End Sub

Public Sub QuickPhaseAcquisition() ' this is similar to the phase acquisition though it does not set the gain and exposure time and triggerMode and focus in advance
    Call triggerMode()
    Call TriggerCamera
    Call TransferFrame
    Call SetAcquisitionEnd
End Sub

Public Sub FlouAcquisition(ByVal picParams As PictureParams)
    
    Call SlowMode()
    Dim focusOffset As Single
    focusOffset = picParams.GetFocusOffset/1000
    If focusOffset <> 0 Then
        stgControl.MoveZBy(focusOffset)
    End If
    Call SetAcquisitionStart(picParams)
    Call triggerMode()
    Call OpenFluoShutter
    Call TriggerCamera
    Call CloseFluoShutter
    Call TransferFrame
    Call SetAcquisitionEnd
    Call FastMode()
    If focusOffset <> 0 Then ' return to focus
        stgControl.MoveZBy(-focusOffset)
    End If
End Sub

Public Sub InitializePreviewMode()
	Call FastMode() 'added 17.5.09
    ret = IpDcamScanMode(DCAM_SCANMODE_FAST)
    ret = IpDcamGain(0)
    moveToFilterNum (red_filter)
End Sub

Public Sub SlowMode()
	ret = IpDcamScanMode(DCAM_SCANMODE_SLOW)
	
End Sub
Public Sub FastMode()
	ret = IpDcamScanMode(DCAM_SCANMODE_FAST)
End Sub
Public Sub InternalMode()
	ret = IpDcamTriggerMode(DCAM_TRIGGER_INTERNAL)
End Sub
Public Sub softwareMode()
	ret = IpDcamTriggerMode(DCAM_TRIGGER_SOFTWARE)
End Sub

Public Sub triggerMode()
	ret=IpDcamStartAcquire(1)   'will start acquisition only upon trigger - added here 18/6/2008
End Sub


Public Sub StartPreview()
	Call InternalMode() 'added 17.5.09
	ret = IpDcamTriggerMode(DCAM_TRIGGER_INTERNAL)
    ret = IpScopeSetPosition(SCP_LIGHT_PATH, 1) 'open bottom port
    Call OpenPhaseShutter
    ret = IpAcqShow(ACQ_LIVE, 1)
End Sub

Public Sub StopPreview()
    ret = IpAcqShow(ACQ_LIVE, 0)
    Call ClosePhaseShutter
    ret = IpDcamTriggerMode(DCAM_TRIGGER_SOFTWARE)
End Sub

Public Sub SetAcquisitionStart(ByVal picParams As PictureParams)
    SetGain (picParams)
    SetExposureTime (picParams)
    moveToFilter(picParams)
    ret = IpMacroWait(1)
    ret = IpDcamTriggerMode(DCAM_TRIGGER_SOFTWARE) 'already set for all movies in main 18/6/2008
    ret = IpDcamStartAcquire(1) 'will start acquisition only upon trigger
End Sub
Private Sub SetAcquisitionEnd()
  ret = IpDcamStopAcquire()
End Sub
Private Sub TriggerCamera()
	ret=IpDcamStartAcquire(1)   'will start acquisition only upon trigger - added here 18/6/2008
    ret = IpDcamFireTrigger()   'trigger to acquire
    ret = IpDcamWaitEvent(DCAM_WAIT_VVALIDBEGIN, timeout)
End Sub
Private Sub TransferFrame()
    ret = IpDcamWaitEvent(DCAM_WAIT_FRAMEEND, timeout)
    ret = IpDcamTransferImages(ACQ_CURRENT) 'transfer frames to current window
End Sub

  



