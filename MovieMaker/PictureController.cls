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
'#Uses "Log.cls"
'Class name : PictureController
'Description : a class controlling the lighting
'Author : Stefan De Prins
'Mail : stefan.de.prins@pandora.be
' needs to pay attention to the filters numbers- is it reflects the microscope settings
Option Explicit
Dim errorLog As New Log
Dim errlogTitle As String
Dim errPathname As String*255
Private exposureTime As Single
Private gain As Integer
Private timeout As Integer
Private stgControl As New StageController
Private empty_filter As Integer
Private focusOffset As Single

Private Sub Class_Initialize()
	errlogTitle = "logging"
	errPathname = "pictureController_errors"&Timer
    gain = 255
    timeout = 20000
    focusOffset = 0.0002
    empty_filter=0
End Sub

Public Sub OpenPhaseShutter()
	LogErr("before open phase shutter ")
    ret = IpScopeControl(SCP_SETCURRSHUTTER, 0, 0, "", IPNULL)
    ret = IpScopeSetPosition(SCP_CURRSHUTTER, 1)
    LogErr("after open phase shutter ")
End Sub

Public Sub ClosePhaseShutter()
	LogErr("before phase shutter close ")
    ret = IpScopeControl(SCP_SETCURRSHUTTER, 0, 0, "", IPNULL)
    ret = IpScopeSetPosition(SCP_CURRSHUTTER, 0)
	LogErr("after phase shutter close ")
End Sub
Public Sub OpenFluoShutter()
	LogErr("before open fluo shutter ")
	ret = IpScopeControl(SCP_SETCURRSHUTTER, 0, 1, "", IPNULL)
    ret = IpScopeSetPosition(SCP_CURRSHUTTER, 1)
    LogErr("before after fluo shutter ")
    End Sub
Public Sub CloseFluoShutter()
	LogErr("before close fluo shutter ")
    ret = IpScopeControl(SCP_SETCURRSHUTTER, 0, 1, "", IPNULL)
    ret = IpScopeSetPosition(SCP_CURRSHUTTER, 0)
	LogErr("after close fluo shutter ")
End Sub

Public Sub moveToFilterNum(ByVal filterNum As Integer)
	LogErr("before move to filter num " & filterNum)
    ret = IpScopeSetPosition(SCP_FWHEEL1, filterNum)
    LogErr("after move to filter num " & filterNum)
End Sub

Public Sub moveToFilter(ByVal picParams As PictureParams)
   LogErr("before move to filter picParams " & picParams.GetFilter())
   Dim filterNum As Integer
   filterNum = picParams.GetFilter
   moveToFilterNum (filterNum)
   LogErr("After move to filter picParams ")
End Sub

Public Sub SetExposureTime(ByVal picParams As PictureParams)
	LogErr("before setExposure time ")
    Dim exposuresec As Single
    exposureTime = picParams.GetExposureTime()
    'Sets the exposure time in mseconds
    If exposureTime < 22 Then
        exposureTime = 22
    End If
    exposuresec = exposureTime / 1000
    IpDcamExposure (exposuresec)
   LogErr("after setExposure time ")
End Sub

Public Sub SetGain(ByVal picParams As PictureParams)
	LogErr("before setGain")
    gain = picParams.GetGain()
    IpAcqControl(53, 1, gain)
    ret = IpDcamGain(gain)
    LogErr("after setGain")
End Sub
Public Sub PicAcquisition(ByVal picParams As PictureParams)
    If picParams.IsPhase Then
        PhaseAcquisition (picParams)
    Else
        FlouAcquisition (picParams)
    End If
End Sub

Public Sub PhaseAcquisition(ByVal picParams As PictureParams)
	Call moveToFilterNum(empty_filter)
	Dim focusOffset As Single
    focusOffset = picParams.GetFocusOffset/1000
 	If focusOffset <> 0 Then
 		LogErr("Beforecall Moving Z")
  	  Call stgControl.MoveZBy(focusOffset)
  	  LogErr("after call Moving Z")
    End If
    '########
		Dim zz As Single
    	 ret = IpStageControl(GETZ, zz)
    	 Debug.Print("phaseAcquisition: taking picture at: "&zz)
    	 '########
    LogErr("Before acquisitionStart")
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
Public Sub QuickPhaseAcquisition(ByVal picParams As PictureParams) ' this is similar to the phase acquisition though it does not set the gain and exposure time and triggerMode and focus in advance
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
    	LogErr("Beforecall Moving Z")
        stgControl.MoveZBy(focusOffset)
    End If
    LogErr("Before acquisitionStart")
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
Public Sub FlouAcquisitionOLD(ByVal picParams As PictureParams)
    ret = IpDcamScanMode(DCAM_SCANMODE_SLOW)
    Dim focusOffset As Single
    focusOffset = picParams.GetFocusOffset/1000
    If focusOffset <> 0 Then
        stgControl.MoveZBy(focusOffset)
    End If
    Call SetAcquisitionStart(picParams)
    Call OpenFluoShutter
    Call TriggerCamera
    Call CloseFluoShutter
    Call TransferFrame
    Call SetAcquisitionEnd
    Call CloseFluoShutter
    ret = IpDcamScanMode(DCAM_SCANMODE_FAST)
    If focusOffset <> 0 Then ' return to focus
        stgControl.MoveZBy(-focusOffset)
    End If
End Sub
Public Sub InitializePreviewMode()
	LogErr("Before init PreviewMode")
    ret = IpDcamScanMode(DCAM_SCANMODE_FAST)
    ret = IpDcamGain(0)
    moveToFilterNum (empty_filter)
    LogErr("After init PreviewMode")
End Sub

Public Sub StartPreview()
	LogErr("Before start PreviewMode")
	Call InternalMode()
    'ret = IpScopeSetPosition(SCP_LIGHT_PATH, 1) 'open bottom port
    Call OpenPhaseShutter
    ret = IpAcqShow(ACQ_LIVE, 1)
    LogErr("After start PreviewMode")
End Sub

Public Sub StopPreview()
	LogErr("Before stop PreviewMode")
    ret = IpAcqShow(ACQ_LIVE, 0)
    Call ClosePhaseShutter
    ret = IpDcamTriggerMode(DCAM_TRIGGER_SOFTWARE)
    LogErr("After stop PreviewMode")
End Sub

Public Sub SetAcquisitionStart(ByVal picParams As PictureParams)
    SetGain (picParams)
    SetExposureTime (picParams)
    moveToFilter(picParams)
    ret = IpMacroWait(1)
    Call softwareMode()
    ret = IpDcamStartAcquire(1) 'will start acquisition only upon trigger
    LogErr("After setAcquisitionStart")
End Sub
Private Sub SetAcquisitionEnd()
	LogErr("Before calling camera to stop acquire")
  ret = IpDcamStopAcquire()
  LogErr("after calling camera to stop acquire")
End Sub
Private Sub TriggerCamera()
	LogErr("Before trigger camera")
	ret=IpDcamStartAcquire(1)   'will start acquisition only upon trigger - added here 18/6/2008
    ret = IpDcamFireTrigger()   'trigger to acquire
    LogErr("Before wait")
    ret = IpDcamWaitEvent(DCAM_WAIT_VVALIDBEGIN, timeout)
    LogErr("After trigger camera")
End Sub
Private Sub TransferFrame()
	LogErr("Before transferFrame")
    ret = IpDcamWaitEvent(DCAM_WAIT_FRAMEEND, timeout)
    Call CloseActiveWindow()
    ret = IpDcamTransferImages(ACQ_CURRENT) 'transfer frames to current window
    LogErr("After transferFrame")
End Sub

  Public Sub SlowMode()
  LogErr("Before slowMode")
	ret = IpDcamScanMode(DCAM_SCANMODE_SLOW)
	LogErr("after slowMode")
End Sub
Public Sub FastMode()
LogErr("Before fastMode")
	ret = IpDcamScanMode(DCAM_SCANMODE_FAST)
	LogErr("After fastMode")
End Sub
Public Sub InternalMode()
LogErr("Before intermalMode")
	ret = IpDcamTriggerMode(DCAM_TRIGGER_INTERNAL)
	LogErr("afterintermalMode")
End Sub
Public Sub softwareMode()
LogErr("Before softwareMode")
	ret = IpDcamTriggerMode(DCAM_TRIGGER_SOFTWARE)
	LogErr("After softwareMode")
End Sub

Public Sub triggerMode()
LogErr("Before triggerMode")
	ret=IpDcamStartAcquire(1)   'will start acquisition only upon trigger - added here 18/6/2008
	LogErr("after triggerMode")
End Sub
Sub CloseActiveWindow
LogErr("Before closeActiveWin")
		ret = IpDocClose() 'closing image to avoid multiple files opened
		LogErr("After closeActiveWin")
End Sub
Sub LogErr(toWrite As String)
    Call errorLog.OpenFile(errPathname, errlogTitle)
    errorLog.WriteLogEntery(" "&toWrite)
    Call errorLog.CloseFile()
End Sub
