'******************************************************************************************************************************
' Name Of Business Component		:	Advanced search - Miscellaneous - Query Instruments
'
' Purpose							:	Advanced search - Miscellaneous - Query Instruments
'
' Input	Parameter					:	
'
' Output							:	True / False
'
' Remarks							:
'
' Author							:	Mohini Deshmukh			12  Nov 2024

'******************************************************************************************************************************
Set objDic=CreateObject("Scripting.dictionary")
'--------------------------------------------------------------------------------------------------------------------------------
'Enter Instrument Number
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_instrument_number")<>""  Then
	objDic.Add glblInstrumentNumber,Parameter("str_instrument_number")
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Select Calibration Database 
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_Calibration_Database") <> "" Then
	objDic.Add gCalibrationDatabase,Parameter("str_Calibration_Database")
End If
''Inactive? ,Calibration Status attributes are removed 12-Nov-24 -Commented By Mohini
''--------------------------------------------------------------------------------------------------------------------------------
''Select Inactive? check box
''--------------------------------------------------------------------------------------------------------------------------------
'If Parameter("str_Calibration_Database") <> "" Then
'	objDic.Add gInactive,Parameter("str_Calibration_Database")
'End If
'--------------------------------------------------------------------------------------------------------------------------------
'Select Responsible Location
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_responsible_location") <>"" Then
	objDic.Add gResponsibleLocation,Parameter("str_responsible_location")
End If
''--------------------------------------------------------------------------------------------------------------------------------
''Select Calibration Status
''--------------------------------------------------------------------------------------------------------------------------------
'If Parameter("str_calibration_status") <>"" Then
'	objDic.Add glblCalibrationStatus,Parameter("str_calibration_status")
'End If
'--------------------------------------------------------------------------------------------------------------------------------
'Enter Calibration Comments
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_calibration_comments") <> "" Then
	objDic.Add glblCalibrationComments,Parameter("str_calibration_comments")
End  If	
'-------------------------------------------------------------------------------------------------------------------------------
'Set Creation Date (from)
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_creation_date_from") <> "" Then
	objDic.Add glblCreationDatefrom2,Parameter("str_creation_date_from")
End If
'-------------------------------------------------------------------------------------------------------------------------------
'Set Creation Date (to)
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_creation_date_to") <> "" Then
	objDic.Add glblCreationDateto2,Parameter("str_creation_date_to")	
End If
'-------------------------------------------------------------------------------------------------------------------------------
'Set Last Modification Date (from)
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_last_modification_date_from") <> "" Then
	objDic.Add gLastModiDatefr,Parameter("str_last_modification_date_from")	
End If
'-------------------------------------------------------------------------------------------------------------------------------
'Set Last Modification Date (to)
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_last_modification_date_to") <> "" Then
	objDic.Add gLastModiDateto,Parameter("str_last_modification_date_to")	
End If
''Calibration Date (from) and Calibration Date (to) attributes are removed 12-Nov-24 -Commented By Mohini
''-------------------------------------------------------------------------------------------------------------------------------
''Set Calibration Date (from)
''--------------------------------------------------------------------------------------------------------------------------------
'If Parameter("str_calibration_date_from") <> "" Then
'	objDic.Add gCalibrationDatefr,Parameter("str_calibration_date_from")	
'End If
''-------------------------------------------------------------------------------------------------------------------------------
''Set Calibration Date (to)
''--------------------------------------------------------------------------------------------------------------------------------
'If Parameter("str_calibration_date_to") <> "" Then
'	objDic.Add gCalibrationDateto,Parameter("str_calibration_date_to")	
'End If
'--------------------------------------------------------------------------------------------------------------------------------
'Enter Name
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_name") <> "" Then
	objDic.Add gName,Parameter("str_name")
End  If	
'--------------------------------------------------------------------------------------------------------------------------------
'Enter Description
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_description") <> "" Then
	objDic.Add gDesc,Parameter("str_description")
End  If	
'--------------------------------------------------------------------------------------------------------------------------------
'Select Owning User
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_Owning_User") <>"" Then
	objDic.Add glblOwningUser,Parameter("str_Owning_User")
End If
'-------------------------------------------------------------------------------------------------------------------------------
'Store search result
'--------------------------------------------------------------------------------------------------------------------------------
Parameter("str_searchresult_out")= Fn_AWC_Advanced_Search_Operations(Parameter("str_expected_search_result"),Parameter("str_searchtype") ,objDic)
objDic.RemoveAll
'-------------------------------------------------------------------------------------------------------------------------------
