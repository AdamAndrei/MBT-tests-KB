'******************************************************************************************************************************
' Name Of Business Component		:	Perform a Do task
'
' Purpose							:	Perform a Do task
'
' Input	Parameter					:	
'
' Output							:	True / False
'
' Remarks							:
'
' Author							:	Mohini Deshmukh 			  18 Dec 2024

'******************************************************************************************************************************
Option Explicit
'-------------------------------------------------------------------------------------------------------------------------------
'Variable Declaration
'--------------------------------------------------------------------------------------------------------------------------------
Dim obj_AWCTeamcenterHome
'--------------------------------------------------------------------------------------------------------------------------------
'Get AWC window object from xml
'--------------------------------------------------------------------------------------------------------------------------------
Set obj_AWCTeamcenterHome=Eval(GetResource("ActiveWorkspace2406_OR.xml").GetValue("wpage_AWCTeamcenterHome"))
''--------------------------------------------------------------------------------------------------------------------------------
Call Fn_AWC_WorkFlow_Perfrom_Task_Operations(obj_AWCTeamcenterHome,Parameter("str_Command"),Parameter("str_Do_Task"),Parameter("str_Workflow_Job_Name"),Parameter("str_addtional_information"))
Call Fn_AWC_ReadyStatusSync(1)
If Parameter("str_WaitTime")<>"" Then
	wait cint(Parameter("str_WaitTime"))
End If
'-----------------------------------------------------------------------------------------------------------------------------------------------------------	
If Err.Number<> 0 Then
	Reporter.ReportEvent micFail, "Perform a Do task", "Fail to perform [ Perform a Do task ]  Operation due to [ " & Err.Description & " ]"
	ExitComponent
Else
	Reporter.ReportEvent micPass, "Perform a Do task", "Successfully performed [ Perform a Do task ] operation "
End If
'-----------------------------------------------------------------------------------------------------------------------------------------------------------	
'Set object nothing
Set obj_AWCTeamcenterHome=Nothing
'-----------------------------------------------------------------------------------------------------------------------------------------------------------


