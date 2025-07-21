'******************************************************************************************************************************
' Name Of Business Component		:	 Create an INSTRUMENT by Tile CREATE
'
' Purpose							:	Perform the Create an INSTRUMENT by Tile CREATE operation
'
' Input	Parameter					:	
'
' Output							:	True / False
'
' Remarks							:
'
' Author							:	Pooja 			  12 may 2022

'******************************************************************************************************************************
Option Explicit
'--------------------------------------------------------------------------------------------------------------------------------
'Variable Declaration
'--------------------------------------------------------------------------------------------------------------------------------
Dim objDic,sOjectNodeName,sObjType,sTemp,sID,sName
Set objDic=CreateObject("Scripting.dictionary")
'--------------------------------------------------------------------------------------------------------------------------------
'Enter Instrument Number
'--------------------------------------------------------------------------------------------------------------------------------
sTemp = Fn_RandNoGenrate(3)
sID = parameter("str_instrument_number")&sTemp
objDic.Add glblInstrumentNumber,sID
Parameter("str_instrument_number_out") = sID
''--------------------------------------------------------------------------------------------------------------------------------
'Enter Name value
'--------------------------------------------------------------------------------------------------------------------------------
sName  = Parameter("str_name")&sTemp
If Parameter("str_name") = "" Then
	Reporter.ReportEvent micFail, "Enter the " &gName& " Value", "Create Business object of type  " & sObjType & " as [ " &gName& " Value is Empty ] "
	ExitComponent
End If
objDic.Add gName2,sName
Parameter("str_name_out") = sName
'--------------------------------------------------------------------------------------------------------------------------------
'Enter Description
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_description") <> "" Then
	objDic.Add gDesc,Parameter("str_description")
	Parameter("str_description_out") =Parameter("str_description") 
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Select Inactive? check box
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_inactive") <> "" Then
	objDic.Add gInactive,Parameter("str_inactive")
	Parameter("str_inactive_out") =Parameter("str_inactive") 
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Select Responsible Location
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_responsible_location") <> "" Then
	objDic.Add gResponsibleLocation,Parameter("str_responsible_location")
	Parameter("str_responsible_location_out") =Parameter("str_responsible_location") 
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Select Calibration Status
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_calibration_database") <> "" Then
	objDic.Add gCalibrationDatabase,Parameter("str_calibration_database")
	Parameter("str_calibration_status_out") =Parameter("str_calibration_database") 
End If
''--------------------------------------------------------------------------------------------------------------------------------
''Enter Calibration Date
''--------------------------------------------------------------------------------------------------------------------------------
'If Parameter("str_calibration_date") <>"" Then
'	objDic.Add glblCalibrationDate,Parameter("str_calibration_date")
'	Parameter("str_calibration_date_out") =Parameter("str_calibration_date") 
'End If
'--------------------------------------------------------------------------------------------------------------------------------
'Enter Calibration Comments
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_calibration_comments") <>"" Then
	objDic.Add glblCalibrationComments,Parameter("str_calibration_comments")
	Parameter("str_calibration_comments_out") =Parameter("str_calibration_comments") 
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Store the search result
'--------------------------------------------------------------------------------------------------------------------------------
sOjectNodeName= Split(Fn_AWC_Object_Creation_Operation("bytile",Parameter("str_type_of_object") ,objDic,""),"~")

Parameter("str_newly_created_object_out")=trim(sOjectNodeName(0))
Parameter("str_instrument_number_out")=trim(sID)
Parameter("str_Full_Node_Name_out")=trim(sID) & ", " & trim(sOjectNodeName(0))
objDic.RemoveAll
'---------------------------------------------------------------------------------------------------------------------------------
