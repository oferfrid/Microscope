VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Location"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Class name :Location
'Description : a container for parameters involving location (currently just a containier for coordinates)
'Author : Stefan De Prins
'Mail : stefan.de.prins@pandora.be

Option Explicit
Private x As Single
Private y As Single
Private z As Single

Public Function GET_X()As Single
    GET_X = x
End Function
Public Sub Set_x(ByVal currX As Single)
    x=currX
End Sub
Public Function Get_y()
    Get_y = y
End Function
Public Sub Set_y(ByVal currY As Single)
    y=currY
End Sub
Public Function GET_Z()
    GET_Z = z
End Function
Public Sub Set_z(ByVal currZ As Single)
    z=currZ
End Sub
