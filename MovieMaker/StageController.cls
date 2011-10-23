VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StageController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Class name : StageController
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

Private stepX As Single
Private stepY As Single
Private stepZ As Single

'Private ret
Public Sub HW()
        MsgBox ("HW")
End Sub

Public Sub MoveTo(ByVal x As Single, ByVal y As Single, ByVal z As Single)
   ret = IpStageXY(x, y)
   ret = IpStageZ(z)
End Sub

Public Sub MoveZ(ByVal z As Single)
   ret = IpStageZ(z)
End Sub

Public Sub MoveXY(ByVal x As Single, ByVal y As Single)
   ret = IpStageXY(x, y)
End Sub

Public Sub MoveX(ByVal x As Single)
        Dim curY As Single
        ret = IpStageControl(GETY, curY)
        ret = IpStageXY(x, curY)
End Sub

Public Sub MoveY(ByVal y As Single)
        Dim curX As Single
   ret = IpStageControl(GETX, curX)
   ret = IpStageXY(curX, y)
End Sub

Public Sub GetCoords(x As Single, y As Single, z As Single)
    x = GetXCoord ()
    y =GetYCoord ()
    z = GetZCoord ()
End Sub

Public Function GetXCoord() As Single
	Dim x As Single
    ret = IpStageControl(GETX, x)
    GetXCoord = x
End Function

Public Function GetYCoord() As Single
	Dim y As Single
    ret = IpStageControl(GETY, y)
    GetYCoord = y
End Function

Public Function GetZCoord() As Single
	Dim z As Single
    ret = IpStageControl(GETZ, z)
    GetZCoord = z
End Function

Public Sub SetStepSizeX(stepSize As Single)
    stepX = stepSize
    ret = IpStageControl(SETSTEPX, stepX)
End Sub

Public Sub SetStepSizeY(stepSize As Single)
    stepY = stepSize
    ret = IpStageControl(SETSTEPY, stepY)
End Sub

Public Sub SetStepSizeZ(stepSize As Single)
    stepZ = stepSize
    ret = IpStageControl(SETSTEPZ, stepZ)
End Sub

Public Sub SetStepSizeAll(stepSize As Single)
    SetStepSizeX (stepSize)
    SetStepSizeY (stepSize)
    SetStepSizeZ (stepSize)
End Sub

Public Sub MoveXBy(ByVal x As Single)
  Dim curX As Single
  curX = GetXCoord()
  Dim nextX As Single
  nextX = curX + x
  MoveX(nextX)
End Sub

Public Sub MoveYBy(ByVal y As Single)
  Dim curY As Single
  curY = GetYCoord()
  Dim nextY As Single
  nextY = curY + y
  MoveY(nextY)
End Sub
Public Sub MoveZBy(ByVal z As Single)
  Dim curz As Single
  curz = GetZCoord()
  Dim nextZ As Single
  nextZ = curz + z
  MoveZ(nextZ)
End Sub

Public Sub MoveOneStepUp()
  IpStageStepZ (STG_UP)
End Sub

Public Sub MoveOneStepDown()
  IpStageStepZ (STG_DOWN)
End Sub
Public Sub MoveOneStepRight()
  IpStageStepXY (STG_RIGHT)
End Sub
Public Sub MoveOneStepLeft()
  IpStageStepXY (STG_LEFT)
End Sub

Public Sub MoveOneStepForward()
  IpStageStepXY (STG_UP)
End Sub

Public Sub MoveOneStepBackward()
  IpStageStepXY (STG_DOWN)
End Sub



