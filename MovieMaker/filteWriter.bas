Option Explicit
Attribute VB_Name = "MicroscopeController"

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



Sub acquisition_bright(gain)
Dim shutter As Integer
    ret=IpDcamGain(gain)
    ret=IpDcamTriggerMode(DCAM_TRIGGER_SOFTWARE)
	ret=IpDcamStartAcquire(1)   'will start acquisition only upon trigger
	open_bright
	ret=IpDcamFireTrigger()	    'trigger to acquire
	ret=IpDcamWaitEvent(DCAM_WAIT_VVALIDBEGIN,20000)
	close_bright

End Sub

Sub runThis()
	Dim i As Integer
	For i = 1 To 100
		Debug.Print "blabla : "&i
		Call acquisition_bright(10)
	Next i
End Sub
Sub open_bright
ret = IpScopeControl(SCP_SETCURRSHUTTER, 0, 0, "", IPNULL)
	ret = IpScopeSetPosition(SCP_CURRSHUTTER, 1)
	End Sub
	Sub close_bright
ret = IpScopeControl(SCP_SETCURRSHUTTER, 0, 0, "", IPNULL)
	ret = IpScopeSetPosition(SCP_CURRSHUTTER, 0)
	End Sub

Sub ooo
	Debug.Print "Ofer : "
End Sub
