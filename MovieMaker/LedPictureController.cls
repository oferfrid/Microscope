VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LedPictureController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'#Reference {244060F0-28B9-3064-A6EC-96375601231E}#2.0#0#C:\Documents and Settings\NQB\Desktop\LED\LedControl\CoolLedControl.tlb#CoolLed Control dll#CoolLedControl
'#Uses "PictureParams.cls"
'#Uses "StageController.cls"
'#Uses "LedFilterChooser.cls"
'#Uses "Log.cls"
'Class name : LedPictureController
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
Private LedFilterNum As Integer
Private ledController As New CoolLedControl.LedControl
Private openShutTime As Single 
Private closeShutTime As Single 
Private LedFilterChooser As New LedFilterChooser


Private Sub Class_Initialize()
	errlogTitle = "logging"
	errPathname = "ledPictureController_errors"
    gain = 255
    timeout = 20000
    empty_filter = 0
    focusOffset = 0.0002
    LedFilterNum = 0
    ret = IpScopeControl(SCP_SETCURRSHUTTER, 0, LedFilterNum, "", IPNULL)
    ret = IpScopeSetPosition(SCP_CURRSHUTTER, 1)
    ledController.openSock()
    Dim i As Integer
    For i= 1 To 3 'set all leds to maximum
    	Call changeIntensity(i,100)
   	Next i
End Sub

Private Sub Class_Terminate()
	LogErr("before treminate ")
	ledController.closeSock()
	ret = IpScopeControl(SCP_SETCURRSHUTTER, 0, LedFilterNum, "", IPNULL)
    ret = IpScopeSetPosition(SCP_CURRSHUTTER, 0)
    LogErr("after treminate ")
End Sub
Public Sub disconnectLed()
LogErr("before disconnect ")
	ledController.closeSock()
	LogErr("after disconnect ")
End Sub

Public Sub OpenPhaseShutter()
LogErr("before openPhaseShutter ")
    ret = IpScopeControl(SCP_SETCURRSHUTTER, 0, 0, "", IPNULL)
    ret = IpScopeSetPosition(SCP_CURRSHUTTER, 1)
    LogErr("after openPhaseShutter ")
End Sub

Public Sub ClosePhaseShutter()
LogErr("before closePhaseShutter ")
    ret = IpScopeControl(SCP_SETCURRSHUTTER, 0, 0, "", IPNULL)
    ret = IpScopeSetPosition(SCP_CURRSHUTTER, 0)
    LogErr("after closePhaseShutter ")
End Sub
Public Sub OpenFluoShutter()
If(LedFilterNum>0 And LedFilterNum<5) Then
	Call openLed()
	End If
	openShutTime =Timer
    End Sub
Public Sub CloseFluoShutter()
If(LedFilterNum>0 And LedFilterNum<5) Then
    Call closeLed()
    End If
    closeShutTime =Timer
End Sub

Public Sub moveToFilterNum(ByVal filterNum As Integer)
LogErr("before moveToFilterNUm " & filterNum)
   ret = IpScopeSetPosition(SCP_FWHEEL1, filterNum)
   LogErr("after moveToFilterNUm " & filterNum)
End Sub
Private Sub openLed()
LogErr("before openLed ")

Debug.Print "from openLed. curFilter is: "& LedFilterNum
	ledController.openLed_2(LedFilterNum,15)
LogErr("after openLed")
End Sub
Private Sub closeLed()
LogErr("before closeLed")
	ledController.closeLed(LedFilterNum)
	LogErr("after closeLed")
End Sub
Public Sub moveToFilter(ByVal picParams As PictureParams)
LogErr("before moveToFilter")
   Dim filterNum As Integer
   filterNum = picParams.GetFilter
   LedFilterNum = LedFilterChooser.getLedFilter(filterNum)
   	moveToFilterNum (filterNum)
   	LogErr("after moveToFilter")
End Sub

Public Sub SetExposureTime(ByVal picParams As PictureParams)
LogErr("before setExpoTime")
    Dim exposuresec As Single
    exposureTime = picParams.GetExposureTime()
    'Sets the exposure time in mseconds
    If exposureTime < 22 Then
        exposureTime = 22
    End If
    exposuresec = exposureTime / 1000
    IpDcamExposure(exposuresec)
   LogErr("after setExpoTime")
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
LogErr("before phaseAcquition")
	Call moveToFilterNum(empty_filter)
	Dim focusOffset As Single
    focusOffset = picParams.GetFocusOffset/1000
 	If focusOffset <> 0 Then
 	LogErr("before  moving Z")
  	  Call stgControl.MoveZBy(focusOffset)
  	  LogErr("after  moving Z")
    End If
    '########
		Dim zz As Single
    	 ret = IpStageControl(GETZ, zz)
    	 Debug.Print("phaseAcquisition: taking picture at: "&zz)
    	 '########
    Call SetAcquisitionStart(picParams)
    Call OpenPhaseShutter
    Call QuickPhaseAcquisition(picParams)
    Call ClosePhaseShutter
    If focusOffset <> 0 Then ' return to focus
        stgControl.MoveZBy(-focusOffset)
    End If
    LogErr("After phaseAcquition")
End Sub
Public Sub QuickPhaseAcquisition(ByVal picParams As PictureParams) ' this is similar to the phase acquisition though it does not set the gain and exposure time and triggerMode and focus in advance
    Call TriggerCamera
    Call TransferFrame
    Call SetAcquisitionEnd
End Sub

Public Sub FlouAcquisition(ByVal picParams As PictureParams)
LogErr("before  FluoAcquisition")
       Call SlowMode()
     Dim focusOffset As Single
    focusOffset = picParams.GetFocusOffset/1000
    If focusOffset <> 0 Then
    LogErr("before  moving Z")
        stgControl.MoveZBy(focusOffset)
        LogErr("after  moving Z")
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
    Debug.Print "opened time : " & openShutTime-closeShutTime
    LogErr("after  FluoAcquisition")
End Sub

Private Sub TransferFrame()
LogErr("before  transferFrame")
    ret = IpDcamWaitEvent(DCAM_WAIT_FRAMEEND, timeout)
    Call CloseActiveWindow()
    ret = IpDcamTransferImages(ACQ_CURRENT) 'transfer frames to current window
    LogErr("After  transferFrame")
End Sub

  Public Sub InitializePreviewMode()
  LogErr("before  InitPrevViewMode")
   Call FastMode()
    ret = IpDcamGain(0)
    moveToFilterNum (empty_filter)
    LogErr("after  InitPrevViewMode")
End Sub
Public Sub SlowMode()
LogErr("before  SlowMode")
	ret = IpDcamScanMode(DCAM_SCANMODE_SLOW)
	LogErr("after  SlowMode")
	
End Sub
Public Sub FastMode()
LogErr("before  fastMode")
	ret = IpDcamScanMode(DCAM_SCANMODE_FAST)
	LogErr("after  fastMode")
End Sub
Public Sub InternalMode()
LogErr("before  internalMode")
	ret = IpDcamTriggerMode(DCAM_TRIGGER_INTERNAL)
	LogErr("after  internalMode")
End Sub
Public Sub softwareMode()
LogErr("before  softwareMode")
	ret = IpDcamTriggerMode(DCAM_TRIGGER_SOFTWARE)
	LogErr("after  softwareMode")
End Sub

Public Sub triggerMode()
LogErr("before  triggerMode")
	ret=IpDcamStartAcquire(1)   'will start acquisition only upon trigger - added here 18/6/2008
	LogErr("after  triggerMode")
End Sub
Public Sub StartPreview()
LogErr("before  startPrev")
	Call InternalMode()
    Call OpenPhaseShutter
    ret = IpAcqShow(ACQ_LIVE, 1)
    LogErr("after  startPrev")
End Sub

Public Sub StopPreview()
LogErr("before  stopPrev")
    ret = IpAcqShow(ACQ_LIVE, 0)
    Call ClosePhaseShutter
    Call softwareMode()
    LogErr("after  stopPrev")
End Sub

Public Sub SetAcquisitionStart(ByVal picParams As PictureParams)
LogErr("before  setAcquisitionStrat")
    SetGain (picParams)
    SetExposureTime (picParams)
    moveToFilter(picParams)
    ret = IpMacroWait(1)
    Call softwareMode()
    ret = IpDcamStartAcquire(1) 'will start acquisition only upon trigger
    LogErr("after  setAcquisitionStrat")
End Sub
Private Sub SetAcquisitionEnd()
LogErr("before  setAcquisitionEnd")
  ret = IpDcamStopAcquire()
  LogErr("after  setAcquisitionEnd")
End Sub
Private Sub TriggerCamera()
LogErr("before  TriggerCam")
	ret = IpDcamFireTrigger()   'trigger to acquire
    ret = IpDcamWaitEvent(DCAM_WAIT_VVALIDBEGIN, timeout)
    LogErr("after  TriggerCam")
End Sub


 Sub CloseActiveWindow
 LogErr("before  closeActWin")
		ret = IpDocClose() 'closing image to avoid multiple files opened
		LogErr("after  closeActWin")
End Sub


Public Sub changeIntensity(ByVal filterNum As Integer,ByVal intens As Integer )
	Dim LedFilterNum As Integer
	LedFilterNum = LedFilterChooser.getLedFilter(filterNum)
	ledController.changeIntensity(LedFilterNum, intens)	
End Sub
Public Sub setPathname(path As String)
	errPathname = IpTrim$(path) &"_ledPictureController_errors"
End Sub

Sub LogErr(toWrite As String)
    Call errorLog.OpenFile(errPathname, errlogTitle)
    errorLog.WriteLogEntery(" Led"&toWrite)
    Call errorLog.CloseFile()
End Sub
Sub setFilterChooser(ByRef fchooser As LedFilterChooser)
fchooser.getLedFilter(1)
	Set LedFilterChooser=fchooser
End Sub
