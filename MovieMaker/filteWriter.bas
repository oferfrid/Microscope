Option Explicit
Attribute VB_Name = "MicroscopeController"

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
