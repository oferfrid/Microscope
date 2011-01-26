VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LocationPics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'#Uses "PictureParams.cls"
'#Uses "Location.cls"

'Class name : LocationPics
'Description : a container for location with the array of picture parameters associated with it

Option Explicit
Private PicParamsArray() As PictureParams
Private mlocation As location
Private aoi As RECT
Private toInitialize As Boolean


Private Sub Class_Initialize()
	toInitialize = True
End Sub
Public Sub SetLocation(Loc As location)
   Set mlocation = Loc
End Sub
Public Sub SetAOI(rectAoi As RECT)
    aoi = rectAoi
End Sub
Public Sub AddPictureType(picParams As PictureParams)
	Dim N As Integer
	If toInitialize Then
		N = 0
		toInitialize=False
	Else
    N = UBound(PicParamsArray)+1
    End If
	ReDim Preserve PicParamsArray(N)
	Set PicParamsArray(N) = New PictureParams
	PicParamsArray(N).copyPictureParams(picParams)
End Sub
Public Sub AddAllPictures(picParams() As PictureParams)
	Dim N As Integer
	N = UBound(picParams())
	Dim i As Integer
	For i=0 To N
		Call AddPictureType(picParams(i))
	Next i
End Sub
Public Function GetLocation() As Location
    Set GetLocation = mlocation
End Function

Public Function GetAOI() As RECT
    GetAOI = aoi
End Function

'#Find Out Why This doesn't work!!!
'Public Function GetPicParamsArray() As PictureParams()
 '  Set GetPicParamsArray = PicParamsArray()
'End Function

Public Function GetPicParams(index As Integer) As PictureParams
    Set GetPicParams= PicParamsArray(index)
End Function

Public Function GetSize() As Integer
	GetSize = UBound(PicParamsArray)
End Function
