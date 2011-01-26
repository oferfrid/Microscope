VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LocationsReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'#Uses "Location.cls"

'Class name : LocationsReader
Dim initExp As Boolean
Dim  lines() As String
Dim locations() As Location
 Sub readFile(pathname As String)
	Dim bytenum As Integer
	Dim size As Integer
	size =0
	Open pathname For Input As #3
    Debug.Print "starting";
    While Not EOF(3)
    	Debug.Print "printing now: " & size
    	ReDim Preserve lines(size) As String
        Line Input #3,lines(size)
        Debug.Print lines(size)
        size=size+1
    Wend
	Debug.Print "========="
    Close #3


End Sub

Function parseLocLine(mLine As String) As Location
	Dim mLoc As New Location
	Dim splited() As String
	Dim splitedX() As String
	Dim splitedy() As String
	Dim splitedZ() As String
	Dim size As Integer
	Call SplitMe(mLine, ":",splited() )
	size = UBound(splited)
	If size=6 Then
		Dim x As Single
		Dim y As Single
		Dim z As Single
		Call SplitMe(splited(3)," ",splitedX())
		Debug.Print "from parseLocLine splitedx(0): " & splitedx(0)
		Debug.Print "from parseLocLine val: " & Val(splitedX(0))
		x =Val(splitedX(0))
		Call SplitMe(splited(4)," ",splitedY())
		Debug.Print "from parseLocLine splitedy(0): " & splitedy(0)
		Debug.Print "from parseLocLine val: " & Val(splitedy(0))
		'splitedy() = Split(splited(4))
		y =Val(splitedy(0))
		'splitedZ() = Split(splited(5))
		Call SplitMe(splited(5)," ",splitedZ())
		Debug.Print "from parseLocLine splitedx(0): " & splitedz(0)
		Debug.Print "from parseLocLine val: " & Val(splitedz(0))
		z =Val(splitedZ(0))
		Debug.Print "from parseLocLine x: " & x & " y: " & y & " z: " & z
		mLoc.Set_x(x)
		mLoc.Set_y(y)
		mLoc.Set_z(z)
	End If
	Set parseLocLine = mLoc
End Function

Private Sub SplitMe(mLine As String, sep As String, splited() As String)
	Dim sepsIndexes() As Integer
	Dim sepNum As Integer
	ReDim Preserve sepsIndexes(sepNum) As Integer
	sepNum = 0
	Dim startFromIndex As Integer
	startFromIndex = 1
	Dim curSepindex As Integer
	Debug.Print "from spliteMe mline: " & mline
	
	curSepIndex = InStr(startFromIndex,mLine,sep) 
	Debug.Print "from spliteMe curSepIndex: " & curSepIndex
	Debug.Print "from spliteMe startFromIndex: " & startFromIndex
	Debug.Print "from spliteMe sep: " & sep
	While curSepIndex>0
		ReDim Preserve sepsIndexes(sepNum) As Integer
		sepsIndexes(sepNum) = curSepIndex
		sepNum=sepNum+1
		startFromIndex = curSepIndex+1
		curSepIndex = InStr(startFromIndex,mLine,sep) 
	Wend
	
	
	Dim i As Integer
	ReDim splited(sepNum+1) As String
	
	Debug.Print "from spliteMe: size is " & (sepNum+1)
	Dim formerInd As Integer
	formerInd = 1
	Debug.Print "spliteMe sepseIndexes size: " & UBound(sepsIndexes())
	For i=0 To (sepNum-1)
		Dim length As Integer
		length = sepsIndexes(i)-formerInd
		'Debug.Print "from spliteMe for i: " & i
		'Debug.Print "from spliteMe for formerInd: " & formerInd
		'Debug.Print "from spliteMe for length: " & length
		'Debug.Print "from spliteMe for mid : " & Mid(mline,formerInd,(length))
		splited(i) = Mid(mline,formerInd+1,length-1)
		Debug.Print "from spliteMe splited(" & i &") : " & splited(i)
		formerInd = sepsIndexes(i)
		
	Next i 
	Debug.Print "from spliteMe last i: " & i
	If i=0 Then 'there are no such seperators
		splited(i) = mLine	
	Else
		splited(i) = Mid(mLine, sepsIndexes(sepNum-1)+1)
		'Debug.Print "from spliteMe last mid : " & Mid(mLine, sepsIndexes(sepNum-1))
	End If
	Debug.Print "from spliteMe splited(" & i &") : " & splited(i)
	Debug.Print "-----------end of spliteMe for i: "
End Sub

Public Sub getLocationsFromFile(path As String, locs() As Location)
	readFile(path)
	Dim i As Integer
	ReDim locations(UBound(lines)) As Location
	For i= 0 To UBound(lines)
		Set locations(i) = parseLocLine(lines(i))
		
	Next i
	'locs() = locations()
	ReDim locs(UBound(locations)) As Location
	For i=0 To UBound(locations)
		Set locs(i) = locations(i)
	Next i
	
End Sub

