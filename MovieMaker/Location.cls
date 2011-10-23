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
' MovieMaker is an ImagePro sofware for time lapsed serial microscope monitoring.
' Copyright 2011 Sivan Pearl published under GPLv3,
' this software was developed in Prof. Nathalie Q. Balaban's lab, at the Hebrew University , Jerusalem , Israel .
'
' This file is part of MovieMaker.
'
' MovieMaker is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' MovieMaker is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with MovieMaker. If not, see <http:'www.gnu.org/licenses/>.


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
