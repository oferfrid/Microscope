Attribute VB_Name = "Module1"
Attribute VB_Name = "stgTest"

Option Explicit

Private stepX As Single

Sub testStage()
stepX = 0.9
ret = IpStageControl(SETSTEPX, stepX)
ret = IpStageStepXY(STG_RIGHT)


	'ret = IpStageXY(0, 0)
	 Debug.Print ret

End Sub

