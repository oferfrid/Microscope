VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LocationsController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'#Uses "StageController.cls"
'#Uses "Location.cls"
'#Uses "Log.cls"
'Class name : LocationsController
'Description : a class controlling the locations of the stage
'Author : Stefan De Prins
'Mail : stefan.de.prins@pandora.be

Option Explicit
Dim logTitle As String
Dim stgControl As New StageController
Private Sub Class_Initialize()
	logTitle = "locations log"
End Sub

Public Function GetCurrLocation() As Location
 	Dim curX As Single
    Dim curY As Single
    Dim curz As Single
    Call stgControl.GetCoords(curX, curY, curz)
    Set GetCurrLocation = GetLocation(curX, curY, curz)
End Function

Public Function GetCurrAOI() As RECT
 	Dim aoiRect As RECT
 	ret = IpTemplateMode(1)
 	ret = IpAoiCreateBox(aoiRect)         'gets aoi from user
 	ret = IpAoiGet(GETBOUNDS, 0, aoiRect)
 	ret = IpTemplateMode(0)
 	ret = IpAoiGet(GETBOUNDS, 0, aoiRect)
	GetCurrAOI = aoiRect
End Function

Public Sub moveToLocation(Location As Location)
	Call stgControl.MoveTo(Location.GET_X, Location.Get_y, Location.GET_Z)
End Sub

Public Sub writeLocation(Location As Location, locLog As Log,ByVal Path As String)
	Call locLog.openFile(Path,logTitle)
    Call locLog.WriteLocLogEntery(Location.GET_X,Location.Get_y,Location.GET_Z)
    Call locLog.CloseFile()
End Sub

Private Function GetLocation(ByVal x As Single, ByVal y As Single, ByVal z As Single) As Location
    Dim Location As New Location
    Location.Set_x (x)
    Location.Set_y (y)
    Location.Set_z (z)
    Set GetLocation = Location
End Function





