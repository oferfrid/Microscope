Option Explicit
Attribute VB_Name = "Module1"
Attribute VB_Name = "MicroscopeController"
'#Uses "PictureParams.cls"
'#Uses "PictureController.cls"
'#Uses "StageController.cls"
'#Uses "LocationsController.cls"
'#Uses "LocationPics.cls"
'#Uses "FocusFinder.cls"
'#Uses "Location.cls"
'#Uses "Log.cls"
'#Uses "LocationsReader.cls"
'Class name : MicroscopeController2
'Description : a class running a microsopy assay
'Author : Tami4

Dim minSecBtwnRounds As Single
Dim stgControl As New StageController
Dim locationsControl As New LocationsController
Dim picControl As New PictureController
Dim phaseExposureTime As Integer
Dim gain As Integer
Dim pathname As String*255
Dim locPathname As String*255
Dim errPathname As String*255
Dim lifeSaverPath As String

Dim timeint As Integer
Dim totRoundsNum As Integer
Dim phaseFocus As New PictureParams
Dim FocusFinder As New FocusFinder
Dim zlogFile As New Log
Dim locLog As New Log
Dim errorLog As New Log

Dim zlogTitle As String
Dim errlogTitle As String
Dim locationsNum As Integer

Dim locations$()
Dim LocationPicsArray() As LocationPics
Dim locIndex As Integer
Dim picTypes$()
Dim picIndex As Integer
Dim PicParamsArray() As PictureParams
Dim initPicParamsArray() As PictureParams
Dim tempPicParams As New PictureParams
Dim defaultPicParams As New PictureParams
Dim initLocations() As location
Dim roundNum As Integer
Dim BaseRoundNum As Integer
Dim focusStep As Single
Dim initRoundTime As Single
Dim endRoundTime As Single
Dim secToWait As Single
Dim initPicArray As Boolean
Dim toInitLocations As Boolean
Dim initExp As Boolean

Sub InitializeParams()
	minSecBtwnRounds=10
	totRoundsNum = 600
	focusStep = 0.001
	timeint = 300
    gain = 255
    roundNum = 0
    zlogTitle = "z coordinates"
    errlogTitle = "logging"
    picIndex = 0
    locIndex = 0
	initPicArray = True
	toInitLocations = True
 	ReDim picTypes$(0)
 	picTypes$(0) = "picture "&picIndex
	ReDim LocationPicsArray(0)
	locationsNum = 0
 	ReDim PicParamsArray(0)
 	ReDim initPicParamsArray(0)
	Call defaultPicParams.SetPhase()
    defaultPicParams.SetFrameFreq (1)
    defaultPicParams.SetExposureTime (40)
    defaultPicParams.SetEmptyFilter
    defaultPicParams.SetGain (255)
    Set PicParamsArray(0) = defaultPicParams
    Set initPicParamsArray(0) = New PictureParams
    initPicParamsArray(0).copyPictureParams(defaultPicParams)
    initExp = False
	inPause = True
End Sub
Sub MovieMaker()

	Call InitializeParams()
    GetParametersFromUser
    picControl.moveToFilter(PicParamsArray(0))
    picControl.InitializePreviewMode
    locationsNum=UBound(LocationPicsArray)
    Call RunPictureManager
	initPicArray = False
	Dim firstLocPic As New LocationPics
	Call firstLocPic.AddAllPictures(initPicParamsArray)
	Set LocationPicsArray(0) = firstLocPic
	While True
    	Call RunLocationManager
    	Call takeRound()
	Wend
End Sub

 'Main loop going over all locations, autofocus, acquiring all type of pictures described in picParamsArray
	Sub takeRound()
		LogErr("startRound "& roundNum)
		initRoundTime  = Timer
		Dim locInd As Integer
		Dim location As location
		Dim LocationPics As LocationPics
		locInd = 0

		Dim endRound As Boolean
	     endRound = False
	     While endRound = False
	      	If locInd = locationsNum Then
	       		endRound = True
		    End If
			Set LocationPics =LocationPicsArray(locInd)
			Set location = LocationPics.GetLocation()
			LogErr("before stage move ")
			locationsControl.moveToLocation(location)
			LogErr("after stage move ")
	       Dim aoirect As RECT
	       aoirect =LocationPics.GetAOI
	       LogErr("startFocusing ")
	       FocusFinder.SET_STEP_Z(focusStep)
	       Call FocusFinder.FindFocus (locInd, aoirect)
	       LogErr("endFocusing ")
	       Call LogCurrZ()
	       'Going over all picParamsArray
	       Dim picParams As PictureParams
		   Dim i As Integer
	       For i = 0 To LocationPics.GetSize
	       	Call updateLifeSaverFile()
	   	    Set picParams = LocationPics.GetPicParams(i)
	   	    Debug.Print "picparams: "& i &" filter: "& picParams.GetFilter()
	   	     If(picParams.shouldTakeFrame(roundNum)) Then
	   	     	LogErr("before pic acquisition ")
		    	 picControl.PicAcquisition (picParams)
		    	 LogErr("after pic acquisition ")
		    	 ret = IpAoiShow(FRAME_NONE)
		    	 If picParams.IsFlour Then
		    	 	ret = IpDrSet(DR_BEST, 0, IPNULL)
		    	 End If
		    	 LogErr("before saving pic ")
	    		 Call saveCurrFile(picParams,pathname, locInd+1, BaseRoundNum+roundNum)
	    		 LogErr("after saving pic ")
	    	 End If
	       Next
	       locInd=locInd+1
		Wend
		LogErr("round finished taking time")
		endRoundTime = Timer
		Debug.Print"finished round: ";roundNum
		roundNum = roundNum + 1
	End Sub

'#################Need to write start/Stop_preview in the PictureController that will not use IpAcqShow, but something from IpDCam...
'################# Log file stuff and Zlog
Sub GetParametersFromUser()
    ret = IpStGetName("Where to save images?", "c:\nathalie\data\temp\", ".tif", pathname)

End Sub
'getting locations from user

Sub getLocationFromUser()
	Dim ans As String*1
    Dim locationNum As Integer
    picControl.StartPreview
    ret = IpMacroStop("move to desire point and acquire", 0)
    picControl.StopPreview
    picControl.PhaseAcquisition (phaseFocus)
    ret = IpDrSet(DR_BEST, 0, IPNULL) 'sets the values for the display range to optimum
    Dim curLoc As location
    Set curLoc = locationsControl.GetCurrLocation()
    Dim currLocPics As LocationPics
    Set currLocPics = LocationPicsArray(locIndex)
    currLocPics.SetLocation(curLoc)
    currLocPics.SetAOI(locationsControl.GetCurrAOI)
    Call currLocPics.AddAllPictures(initPicParamsArray())
    FocusFinder.AddLocation(curLoc)

    locPathname = IpTrim$(pathname) & "_locations"
    errPathname = IpTrim$(pathname) &"_errors"
    Debug.Print "path is: " & locPathname
    Call locationsControl.writeLocation(curLoc, locLog,locPathname)
End Sub
'getting locations from a file, than marking aois.
Sub addLocationsFromFile()
	Dim  fileLocations() As location
	Dim ans As String*1
    Dim locationNum As Integer
    Dim reader As New LocationsReader

    ret = IpStGetName("locations file?", "c:\nathalie\data\temp\", "", locPathname)
    'locPathname = "C:\Documents and Settings\Owner\Desktop\loc"
	Call reader.getLocationsFromFile(locPathname,fileLocations)
	Dim locNum As Integer
	locNum = UBound(fileLocations)
	Dim i As Integer
	For i=0 To locNum
		Dim location As location
		Set location = fileLocations(i)
		locationsControl.moveToLocation(location)
		Call AddLocation()
 	Next i
End Sub
'writes to the z log file the current time and z coordinates.
Sub LogCurrZ()
    Dim currZ As Single
    Call zlogFile.OpenFile(pathname, zlogTitle)
    currZ = stgControl.GetZCoord()
    zlogFile.WriteZLogEntery(currZ)
    Call zlogFile.CloseFile()
End Sub
Sub LogErr(toWrite As String)
    Call errorLog.OpenFile(errPathname, errlogTitle)
    errorLog.WriteLogEntery(toWrite)
    Call errorLog.CloseFile()
End Sub
Sub saveCurrFile(picParams As PictureParams, pathname As String,locationNum As Integer, roundNum As Integer)
	Dim fname As String
	fname = IpTrim$(pathname)&"_"&picParams.GetPicName()&"_"&locationNum &"_"&roundNum+1 &".tif"
	ret=IpWsSaveAs(fname, "tif")
End Sub

Sub RunLocationManager
  '################CHanged
	Begin Dialog UserDialog 370,294,.locationFunc ' %GRID:10,7,1,1
		TextBox 240,210,110,21,.txtRoundsNum
		Text 20,245,200,14,"focus Step",.Text3
		Text 20,273,200,14,"starting round number",.Text4
		ListBox 10,25,180,60,locations(),.locations
		PushButton 20,98,70,21,"Edit",.btnEditLoc
		PushButton 130,98,60,21,"Add",.btnAddLoc
		PushButton 200,49,150,28,"Add locations from file",.btnAddFileLoc
		PushButton 30,133,150,21,"Cont.",.btnStartPause
		TextBox 240,168,110,21,.txtTimeInt
		Text 10,168,220,21,"Seconds from 1 round to another:",.Text1
		Text 20,210,200,14,"rounds number:",.Text2
		TextBox 240,245,110,21,.txtFocusStep
		TextBox 240,273,110,21,.txtBaseRoundNum
		PushButton 280,7,80,28,"Close",.btnClose
	End Dialog
	Dim dlg As UserDialog
	dlg.txtTimeInt = CStr(timeint)
	dlg.txtRoundsNum = CStr(totRoundsNum)
	dlg.txtFocusStep = CStr(focusStep)
	dlg.txtBaseRoundNum = CStr(BaseRoundNum)

	Dim ans As Integer
    Dialog dlg
End Sub
Rem See DialogFunc help topic for more information.
Dim inPause As Boolean
Dim printTimeInterval As Single
Dim prevSecToWait As Single

Private Function LocationFunc%(DlgItem$, Action%, SuppValue?)
	Select Case Action%
	Case 1 ' Dialog box initialization
		printTimeInterval = 15
		prevSecToWait = 0
		If inPause Then
			DlgText "btnStartPause", "cont."
		Else
			DlgText "btnStartPause", "pause"

		End If
	Case 2
	Select Case DlgItem$

		Case "btnClose"
			ret=IpMacroStop("Close program?",MS_MODAL+MS_YESNO+ MS_QUEST)
			If ret = 1 Then
				End
			Else
			 	LocationFunc%= True 'continue program
			End If

		Case "btnEditLoc"
			inPause = True
			Call RunPictureManager
			LocationFunc% = True 'do not exit the dialog
			inPause = True 'Changed 17 Aug2010
			DlgText "btnStartPause", "cont."
		Case "btnAddLoc"
			inPause = True
			Call AddLocation()
		    LocationFunc% = True 'do not exit the dialog
		    DlgText "btnStartPause", "cont."
		Case "btnAddFileLoc"
			inPause = True
			DlgText "btnStartPause", "cont."
			Call addLocationsFromFile()
		    LocationFunc% = True 'do not exit the dialog
		Case "btnStartPause"
			If initExp = False Then
				initExp = True
			End If
			If inPause = True Then
				inPause = False
			    DlgText "btnStartPause", "Pause"
			Else
				inPause = True
				DlgText "btnStartPause", "Cont."
			End If
			LocationFunc% = True 'do not exit the dialog

		Case "locations"
			locIndex =SuppValue


End Select
	' Value changing or button pressed
		  '################CHanged
	Case 3 ' TextBox or ComboBox text changed
		Select Case DlgItem$
			Case "txtTimeInt"
				timeint = DlgText("txtTimeInt")
			Case "txtRoundsNum"
				totRoundsNum = DlgText("txtRoundsNum")
			Case "txtFocusStep"
				focusStep = DlgText("txtFocusStep")
			Case "txtBaseRoundNum"
				BaseRoundNum = DlgText("txtBaseRoundNum")
		End Select
	Case 5 ' Idle
		LocationFunc = True ' Continue getting idle actions
		' Clear the event queue, so that we can respond to buttons
		' and other events
		DoEvents
		If initExp = True Then
			secToWait = getSecToWaitWithMinimum(minSecBtwnRounds)
			If((Abs(secToWait-prevSecToWait))>printTimeInterval) Then
				Debug.Print("seconds till next round: "&secToWait)
				prevSecToWait = secToWait
			End If
			If(roundNum>=totRoundsNum) Then
				IpMacroStop("finished "& roundNum &" Rounds",0)
				ret=IpMacroStop("finished "& roundNum &" Rounds",MS_MODAL+MS_YESNO+ MS_QUEST)
				If ret = 1 Then
					End
				Else
			 		LocationFunc%= True 'continue program
				End If
			End If

			If (inPause = False) And (secToWait<0) Then
			    DlgEnd 1
   			End If
   		End If
	End Select
End Function
Sub AddLocation()
	Dim N As Integer
	If toInitLocations = True Then
		N=0
		toInitLocations = False
	Else
	    N = UBound(locations$)+1
	End If
    ReDim Preserve locations$(N)
	ReDim Preserve LocationPicsArray(N)
	ReDim Preserve LocationsArray(N)
    locations$(N) = "Location " & N
    DlgListBoxArray "Locations",locations$()
	Set LocationsArray(N) = New location
	Set LocationPicsArray(N) = New LocationPics
	locIndex = N
 	locationsNum =N
	Call getLocationFromUser()

End Sub
Sub RunPictureManager
	Call createPicTypeList()
	Begin Dialog UserDialog 370,245,.picParamsFunc ' %GRID:10,7,1,1
		ListBox 10,25,180,60,picTypes(),.list
		PushButton 20,98,70,21,"Edit",.btnEdit
		PushButton 110,98,60,21,"Add",.btnAdd
		OKButton 280,14,60,21
	End Dialog
	Dim dlg As UserDialog
	Dim ans As Integer
    Dialog dlg
End Sub

Sub EditPic()
	If initPicArray Then
		tempPicParams.copyPictureParams(initPicParamsArray(picIndex))
	Else
		Dim currLocPic As LocationPics
		Set currLocPic = LocationPicsArray(locIndex)
		tempPicParams.copyPictureParams(currLocPic.GetPicParams(picIndex))
	End If
	Begin Dialog UserDialog 480,364,"Edit Picture Parameters",.EditFunc ' %GRID:10,7,1,1
		TextBox 140,273,210,21,.txtFreq
		Text 20,238,90,14,"offset (microns)",.Text5
		TextBox 140,238,210,21,.txtOffset
		Text 20,273,90,14,"Frequency",.Text4
		TextBox 140,203,210,21,.txtGain
		Text 20,168,90,14,"Exposure(sec)",.Text3
		TextBox 130,14,200,21,.picName
		TextBox 140,168,210,21,.txtExposureTime
		OKButton 360,329,110,21
		Text 20,21,90,14,"Name",.Text1
		Text 20,203,90,14,"gain:",.Text2
		OptionGroup .GroupType
			OptionButton 10,63,100,14,"Phase",.btnPhase
			OptionButton 10,91,120,14,"Fluoresence",.btnFluo
		GroupBox 170,49,250,91,"Choose Filter",.FilterGroup
		OptionGroup .GroupFilters
			OptionButton 180,77,110,21,"0",.optEmpty_F
			OptionButton 180,105,110,21,"1",.OptCherry_F
			OptionButton 280,77,120,21,"2",.OptKO_F
			OptionButton 300,105,110,21,"3",.optYFP_F
		Text 20,294,330,14,"1 means every round, 2 every second round etc.",.Text6
	End Dialog
	Dim dlg As UserDialog
	dlg.picName = "picture "&picIndex
	dlg.txtExposureTime = CStr(tempPicParams.GetExposureTime)
	dlg.txtGain = CStr(tempPicParams.GetGain)
	dlg.txtOffset = CStr(tempPicParams.GetFocusOffset)
	dlg.txtFreq = CStr(tempPicParams.GetFrameFreq)
	If tempPicParams.IsPhase  = True Then
		dlg.GroupType = 0
	Else
	  	dlg.GroupType = 1
	End If
	dlg.GroupFilters = tempPicParams.GetFilter
	Dialog dlg
End Sub

Sub AddPic()
	Dim N As Integer
	If initPicArray Then
	    N = UBound(initPicParamsArray)+1
		ReDim Preserve initPicParamsArray(N)
		Set initPicParamsArray(N) = New PictureParams
		initPicParamsArray(N).copyPictureParams(defaultPicParams)
	Else
		Dim currLocPic As LocationPics
		Set currLocPic = LocationPicsArray(locIndex)
	 	N = currLocPic.GetSize+1
		currLocPic.AddPictureType(defaultPicParams)
	End If

	Call createPicTypeList()
    DlgListBoxArray "list",picTypes$()
End Sub

Rem See DialogFunc help topic for more information.

Private Function picParamsFunc%(DlgItem$, Action%, SuppValue?)

	Select Case Action%
	Case 1 ' Dialog box initialization

	Case 2

	Select Case DlgItem$
		Case "OK"

		Case "btnEdit"
			'inEditMode = True
			inPause = True
			Call EditPic()
			picParamsFunc% = True 'do not exit the dialog
			'inEditMode = False
			inPause = True
		Case "btnAdd"
			Call AddPic()
		    picParamsFunc% = True 'do not exit the dialog

		Case "list"
			picIndex =SuppValue
End Select

	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle

	Case 6 ' Function key
	End Select

End Function


Rem See DialogFunc help topic for more information.
Private Function EditFunc(DlgItem$, Action%, SuppValue?) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
	Case 2 ' Value changing or button pressed
		If DlgItem$ = "OK" Then 'ok was clicked
			tempPicParams.SetExposureTime(DlgText("txtExposureTime"))
			tempPicParams.SetFocusOffset(DlgText("txtOffset"))
			tempPicParams.SetFrameFreq(DlgText("txtFreq"))
			tempPicParams.SetGain(DlgText("txtGain"))
			Dim optBtnVal As Integer
			optBtnVal = DlgValue("GroupType")
			If optBtnVal = 0 Then
				tempPicParams.SetPhase
			ElseIf optBtnVal = 1 Then
				tempPicParams.SetFluo
			End If
			optBtnVal = DlgValue("GroupFilters")
			Select Case optBtnVal
				Case 0
					tempPicParams.SetFilter(0)
				Case 1
					tempPicParams.SetFilter(1)
				Case 2
					tempPicParams.SetFilter(2)
				Case 3
					tempPicParams.SetFilter(3)
			End Select
		End If
		If initPicArray Then
			initPicParamsArray(picIndex).copyPictureParams(tempPicParams)
		Else
			LocationPicsArray(locIndex).GetPicParams(picIndex).copyPictureParams(tempPicParams)
		End If

		Rem EditFunc = True ' Prevent button press from closing the dialog box
	Case 3 ' TextBox or ComboBox text changed
	Case 4 'Focus changed
	Case 5 ' Idle

		Rem Wait .1 : EditFunc = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function

Function getSecToWaitWithMinimum(minimumWait As Single) As Single
	Dim ans As Single
	Dim currTime As Single
	Dim timeSinceEndRound As Single
	currTime = Timer
	'need to wait minimum time
	timeSinceEndRound =getDeltaSecs(endRoundTime,currTime)
	If timeSinceEndRound<minimumWait Then
		ans = minimumWait - timeSinceEndRound
	Else 'finished minimum time wait now we wait for the timeInterval between rounds
		Dim timeSinceStart As Single
		timeSinceStart = getDeltaSecs(initRoundTime,currTime)
		ans = timeint - timeSinceStart
	End If
	getSecToWaitWithMinimum = ans
End Function
Function getDeltaSecs(a As Single, b As Single) As Single
	If(a>b) Then
		getDeltaSecs =86400-a+b
	Else
		getDeltaSecs = b-a
	End If
End Function

Sub createPicTypeList()
	Dim N As Integer
	If initPicArray Then
	    N = UBound(initPicParamsArray)
	Else
		Dim currLocPic As LocationPics
		Set currLocPic = LocationPicsArray(locIndex)
	 	N = currLocPic.GetSize
	End If
	Dim i As Integer
	ReDim picTypes$(N)
 	For i = 0 To N Step 1
		picTypes$(i) ="Picture " & i
		Debug.Print "add picture: "& picTypes$(i)
    Next i
End Sub

Sub setInitLocations()
	Dim N As Integer
	N = UBound(LocationPicsArray)
	ReDim initLocations(N)
	Dim i As Integer
	For i=0 To N
		Set initLocations(i) = LocationPicsArray(i).GetLocation
	Next i

End Sub
'writes the current time to the lifeSaverUpdate file, overriding the current file
Sub updateLifeSaverFile()
	Dim lifeSaverPath As String
	lifeSaverPath = "C:\Documents and Settings\NQB\Desktop\microscope scripts\microscope scripts LED 1.1\lifeSaverIproUpdate.txt"
	Dim iFileNo As Integer
	iFileNo = FreeFile
	Debug.Print "lifeSaverPath is: " & lifeSaverPath
	Open lifeSaverPath For Output As iFileNo
	Print #iFileNo, Now()
	Close #iFileNo
End Sub
