VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FocusFinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'#Uses "StageController.cls"
'#Uses "PictureController.cls"
'#Uses "Location.cls"
'#Uses "PictureParams.cls"
'Class name : FocusFinder
'Description : a class controlling the Focus
'Author : Stefan De Prins
'Mail : stefan.de.prins@pandora.be
Option Explicit
Private stgControl As New StageController
Private stepZ As Single 'stepZ is the step size made by the stage when scanning for best focus

Private frameNum As Integer 'framenum is the number of pics taken to determine the best focus.

Private picControl As New PictureController
Private exposureTime As Single
Private unintializedZ As Single

Private locationsNum As Integer
Private stds() As Single
Private badFocuses() As Single
Private bestZs() As Single
'Defining the type of phase picture this picController will take for the finf focus sub
Private picParams As New PictureParams

'-----------
Private Sub Class_Initialize()
    stepZ = 0.001
    frameNum = 20
    exposureTime = 40
    picParams.SetFrameFreq (1)
    picParams.SetExposureTime (exposureTime)
    picParams.SetEmptyFilter
    picParams.SetGain (100)
    unintializedZ = 10000
    locationsNum=0

End Sub
Public Sub FindFocus(ByVal locationNum As Integer, aoirect As RECT )
    Dim stdmax As Single
    Dim stdcurr As Single
    Dim framebest As Integer
    Dim stats(10) As Single
    Dim bestz As Single
    ReDim stddev(frameNum) As Single
    Dim zcoord As Single 'for debugging

    MoveToLastZ (locationNum)
    ipStgVal = stepZ * frameNum * (0.55 - badFocuses(locationNum) * 0.35) 'goes down a bit more than up ; if focus bad (pp.badfocus=1), goes mainly up
    zcoord = stgControl.GetZCoord()
    stgControl.MoveZBy(-ipStgVal) 'goes down first by logical steps ipStgVal.
    zcoord = stgControl.GetZCoord()
    Debug.Print "Down at first stageZ:"; zcoord; " next step down:"; stepZ
    framebest = 0
    stdmax = 0
    stdcurr = 60000

    picControl.SetAcquisitionStart (picParams)
    picControl.OpenPhaseShutter
	Dim i As Integer
    For i = 1 To frameNum
		stgControl.MoveZBy(stepZ)
    	zcoord = stgControl.GetZCoord()
    	Debug.Print i; " z "; zcoord; "  ";
    	picControl.QuickPhaseAcquisition()
    	If(i=1) Then
			ret = IpDrSet(DR_RANGE, -1, IPNULL) 'get all range display
		End If

    	ret = IpAoiCreateBox(aoirect)
    	ret = IpFltVariance(5, 5)                'filter to detect and emphasize edges and textures
    	ret = IpHstGet(GETSTATS, 0, stats(0))    'gets the histogram parameters of the aoi stats(0) is mean and stats(1) is std
    	stddev(i) = stats(1)               'compares standart deviation of the frames
    	Debug.Print "stddev"; stddev(i); "  ";
    	If stddev(i) > stdmax Then
      		stdmax = stddev(i)           'searches for maximal std deviation
      		framebest = i
      		bestz = zcoord
    	End If
    
    	Debug.Print "best "; framebest; " maxstddev "; stdmax
    Next i
    picControl.ClosePhaseShutter

    stdcurr = stdmax
    If (framebest < 2 Or framebest > (frameNum - 2)) Then 'defines problematic focus if the best frames are chosen from the extremeties of focus search
     badFocuses(locationNum) = 1
     Debug.Print "Bad focus?? "
    Else
	 badFocuses(locationNum) = 0
    End If
    bestZs(locationNum) = bestz
    Debug.Print "best z for location: "; locationNum; " is :" ; bestZs(locationNum)
    stds(locationNum) = stdmax
    stgControl.MoveZ(bestz)

End Sub
Public Sub initBestLocations(initLocations() As location, N As Integer)
	Dim i As Integer
	Dim i2 As Integer
	i2 = LBound(bestZs)
	For i=LBound(initLocations) To N
		If IsEmpty(initLocations(i)) = False Then
			Debug.Print "in initBestLocations i: ";i;" out of: ";N;" initlocation: ";initLocations(i).get_Z
			bestZs(i2) = initLocations(i).get_Z
			Debug.Print "in initBestLocations bestZs i2: ";i2;" is: ";bestZs(i2)
		End If
		i2=i2+1
	Next i

End Sub
Public Sub AddLocation(location As location)
	Dim N As Integer
	If locationsNum = 0 Then
		N=0
	Else
		N = UBound(bestZs)+1
	End If
	ReDim Preserve bestZs(N) As Single
	ReDim Preserve badFocuses(N) As Single
	ReDim Preserve stds(N) As Single
	bestZs(N) = location.get_Z()
	locationsNum = locationsNum+1
End Sub

Public Sub SET_STEP_Z(ByVal Step As Single)
    stepZ = Step
End Sub

Public Sub SetFrameNum(ByVal N As Single)
    frameNum = N
End Sub

Public Sub SetExposureTime(ByVal Time As Single)
    exposureTime = Time
    picParams.SetExposureTime (Time)
End Sub

Private Sub MoveToLastZ(ByVal locationNum As Integer)  'moves to last z of this locationNum only if it had a good focus last time.
		Dim z As Single
		z = stgControl.GetZCoord()
		Debug.Print"before moveToLastZ : ";z
		If bestZs(locationNum) <> unintializedZ Then
       		 stgControl.MoveZ (bestZs(locationNum))
       		 Debug.Print"inside moveToLastZ bestZs: ";bestZs(locationNum)
        End If
        	z = stgControl.GetZCoord()
        Debug.Print"after moveToLastZ : "; z
End Sub
