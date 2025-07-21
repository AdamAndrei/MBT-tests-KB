'******************************************************************************************************************************
' Name Of Business Component		:	04.02.00.09. Add attachment to INSTRUMENTS by Paste (Add the content of the clipboard here)
'
' Purpose							:	04.02.00.09. Add attachment to INSTRUMENTS by Paste (Add the content of the clipboard here)
'
' Input	Parameter					:	
'
' Output							:	True / False
'
' Remarks							:
'
' Author							:	Mohini Deshmukh 			  27 Dec 2024

'******************************************************************************************************************************
Option Explicit
'-------------------------------------------------------------------------------------------------------------------------------
'Variable Declaration
'--------------------------------------------------------------------------------------------------------------------------------
Dim sTemp,objDic
Set objDic=CreateObject("Scripting.dictionary")
'--------------------------------------------------------------------------------------------------------------------------------
Call Fn_AWC_Add_Attachment_By_Paste(Parameter("str_primary_object") ,Parameter("str_copied_object") ,gtabRelations,gINSTRUMENTS,"")
Parameter("str_copied_object_out") = Parameter("str_copied_object")
Parameter("str_primary_object_out") = Parameter("str_primary_object")
'-----------------------------------------------------------------------------------------------------------------------------------------------------------	
objDic.RemoveAll
'-----------------------------------------------------------------------------------------------------------------------------------------------------------	
	
