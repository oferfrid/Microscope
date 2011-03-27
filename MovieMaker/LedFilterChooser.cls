VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LedFilterChooser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Class name : LedFilterChooser
'Description : a container for the association of filter number with type of LED
'Author : Stefan De Prins
'Mail : stefan.de.prins@pandora.be

Option Explicit
 Private filter1 As Integer
Private filter2 As Integer
Private filter3 As Integer

'filter is the filter of the microscope the number is of the led to operate
Private Sub Class_Initialize()
	filter1  = 0 'supposed to be empty
	filter1  = 3 'usually  for Cherry/KO
	filter2 = 2 'change according to what is needed
	filter3  = 2 'usually for GFP/YFP
End Sub
Public Function getLedFilter(micFilter As Integer) As Integer
    Dim ledFilter As Integer
    Select Case micFilter
    		Case 0
    			ledFilter = filter0
			Case 1
				ledFilter = filter1
			Case 2
				ledFilter = filter2
			Case 3
				ledFilter = filter3
		End Select
    getLedFilter = ledFilter
End Function

Public Function setLedFilter(micFilter As Integer,ledFilter As Integer)
		Select Case micFilter
			Case 0
				filter0=ledFilter
			Case 1
				filter1=ledFilter
			Case 2
				filter2 = ledFilter
			Case 3
				filter3 = ledFilter
		End Select
End Function






