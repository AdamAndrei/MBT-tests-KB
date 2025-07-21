'******************************************************************************************************************************
' Name Of Business Component		:	Alternate Work
'
' Purpose							:	Alternate Work
'
' Input	Parameter					:	
'
' Output								:	True / False
'
' Remarks							:
'
' Author								:	Mohini Deshmukh			  16 Dec 2024

'******************************************************************************************************************************
Option Explicit
'-------------------------------------------------------------------------------------------------------------------------------
'Variable Declaration
'--------------------------------------------------------------------------------------------------------------------------------
Dim obj_AWCTeamcenterHome
'--------------------------------------------------------------------------------------------------------------------------------
'Get AWC  window object from xml
'--------------------------------------------------------------------------------------------------------------------------------
Set obj_AWCTeamcenterHome=Eval(GetResource("ActiveWorkspace2406_OR.xml").GetValue("wpage_AWCTeamcenterHome"))
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_AlternateFlag")<>"False" Then
	'-----------------------------------------------------------------------------------------------------------------------------------------------------------
	If Parameter("str_Performance_Monitor")="True" Then	
		'Perform alternate work task operation
		Call Fn_AWC_WorkFlow_AlternateWork_Operations("selectrow~Performance~AWC2406_Alternate Work.xlsx",obj_AWCTeamcenterHome, Parameter("str_RecepientUser") ,Parameter("str_SectionName"),Parameter("str_WorkFlowJobName"),Parameter("str_TaskName"),Parameter("str_ReassignedUser"),Parameter("str_Comment"))
	Else
		'Perform alternate work task operation
		Call Fn_AWC_WorkFlow_AlternateWork_Operations("selectrow",obj_AWCTeamcenterHome, Parameter("str_RecepientUser") ,Parameter("str_SectionName"),Parameter("str_WorkFlowJobName"),Parameter("str_TaskName"),Parameter("str_ReassignedUser"),Parameter("str_Comment"))
	End IF	
	'-----------------------------------------------------------------------------------------------------------------------------------------------------------
Else
	Parameter("str_ReassignedUser_out")=Parameter("str_RecepientUser")
End If

Parameter("str_Comment_out")=Parameter("str_Comment")
'---------------------------------------------------------------------------------------------------------------------------------
If Err.Number<> 0 Then
	Reporter.ReportEvent micFail, "Alternate Work", "Fail to perform [ Alternate Work ]  Operation due to [ " & Err.Description & " ]"
	ExitComponent
Else
	Reporter.ReportEvent micPass, "Alternate Work", "Successfully performed [ Alternate Work ] operation "
End If
'-----------------------------------------------------------------------------------------------------------------------------------------------------------	
'Release the object 
Set obj_AWCTeamcenterHome=Nothing
'-----------------------------------------------------------------------------------------------------------------------------------------------------------


