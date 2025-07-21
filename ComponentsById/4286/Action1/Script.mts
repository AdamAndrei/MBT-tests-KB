'******************************************************************************************************************************
' Name Of Business Component		:	Add attachments -  Files - Other by Choose File -  to an item revision at tab Overview
'
' Purpose							:	Add attachments -  Files - Other by Choose File -  to an item revision at tab Overview
'
' Input	Parameter					:	
'
' Output							:	True / False
'
' Remarks							:
'
' Author							:	Mohini Deshmukh		  21 March 2022

'******************************************************************************************************************************
Option Explicit
'--------------------------------------------------------------------------------------------------------------------------------
'Variable Declaration
'--------------------------------------------------------------------------------------------------------------------------------
Dim sTemp,objDic
Set objDic=CreateObject("Scripting.dictionary")
'--------------------------------------------------------------------------------------------------------------------------------
sTemp= Fn_AWC_Add_File_By_Choose_File(Parameter("str_ObjectID"),gtabOverview,Parameter("str_TabName"),Parameter("str_filepath"),Parameter("str_filename"),Parameter("str_filedescription"),objDic)
Parameter("str_file_name_out") = sTemp
objDic.RemoveAll
'-----------------------------------------------------------------------------------------------------------------------------------------------------------


