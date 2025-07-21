'******************************************************************************************************************************
' Name Of Business Component		:	Validate  Primary toolbar  - Copy
'
' Purpose							:	Validate  Primary toolbar  - Copy
'
' Input	Parameter					:	
'
' Output							:	True / False
'
' Remarks							:
'
' Author							:	Mohini Deshmukh		11 April 2022

'******************************************************************************************************************************
Option Explicit
'-------------------------------------------------------------------------------------------------------------------------------
'Variable Declaration
'--------------------------------------------------------------------------------------------------------------------------------
Dim obj_AWCTeamcenterHome,sStatusMesaage
'--------------------------------------------------------------------------------------------------------------------------------
'Get AWC window object from xml
'--------------------------------------------------------------------------------------------------------------------------------
Set obj_AWCTeamcenterHome=Eval(GetResource("ActiveWorkspace2406_OR.xml").GetValue("wpage_AWCTeamcenterHome"))
'--------------------------------------------------------------------------------------------------------------------------------
'Select the Object
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_ObjectName")<>"" Then
	Call Fn_AWC_Object_Navigation_Operations(obj_AWCTeamcenterHome,"select",Parameter("str_ObjectName") )
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Click on New Toolbar button on Primary tab
'--------------------------------------------------------------------------------------------------------------------------------
If Fn_WEB_UI_WebButton_Operations("Validate  Primary toolbar  - Copy", "Click", obj_AWCTeamcenterHome, "wbtn_Copy","","","") =False Then
	Reporter.ReportEvent micFail, "Click on [ Copy ] toolbar button on primary toolbar", "Fail to click on [ Copy ] toolbar button on primary toolbar"
	ExitComponent
Else
	Reporter.ReportEvent micPass, "Click on [ Copy ] toolbar button on primary toolbar", "Successfully clicked on [ Copy ] toolbar button on primary toolbar"
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Verify notification message
'--------------------------------------------------------------------------------------------------------------------------------
If  WaitUntilExist(obj_AWCTeamcenterHome.WebElement("wele_ObjectCreationNotificationMsg"), 5, 30) Then  
	sStatusMesaage= Fn_Web_UI_WebObject_Operations("Validate  Primary toolbar  - Copy", "getroproperty", obj_AWCTeamcenterHome.WebElement("wele_ObjectCreationNotificationMsg"), "2", "innertext", "")
	If  instr(sStatusMesaage, "copied to Teamcenter and OS clipboard" )>0 Then
		Reporter.ReportEvent micPass, "Verify Notification message", "Successfully verified notification message which is [ "&  sStatusMesaage  & " ]"
	Else
		Reporter.ReportEvent micFail, "Verify Notification message", "Fail to copy object ["&Parameter("str_ObjectName") &"] as [ Fail to verify notification message ]"
		ExitComponent
	End If	
End  IF

Parameter("str_ObjectName_out") = Parameter("str_ObjectName")
'--------------------------------------------------------------------------------------------------------------------------------
If Err.Number<> 0 Then
	Reporter.ReportEvent micFail, "Validate  Primary toolbar  - Copy", "Fail to perform [ Validate  Primary toolbar  - Copy ]  Operation due to [ " & Err.Description & " ]"
	ExitComponent
Else
	Reporter.ReportEvent micPass, "Validate  Primary toolbar  - Copy", "Successfully performed [ Validate  Primary toolbar  - Copy ] operation "
End If
'-----------------------------------------------------------------------------------------------------------------------------------------------------------	
'Release the object 
Set obj_AWCTeamcenterHome=Nothing
'-----------------------------------------------------------------------------------------------------------------------------------------------------------	
