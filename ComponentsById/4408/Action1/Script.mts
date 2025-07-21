'******************************************************************************************************************************																																																																																																																																																																																																																																																				'******************************************************************************************************************************
' Name Of Business Component		:	Signout Active Workspace
'
' Purpose							:	Signout the Active Workspace application
'
' Input	Parameter					:	
'
' Output							:	True / False
'
' Remarks							:
'
' Author							:	Mohini Deshmukh 			  1 Aug 2020

'******************************************************************************************************************************
'-------------------------------------------------------------------------------------------------------------------------------
'Variable Declaration
'-------------------------------------------------------------------------------------------------------------------------------
Dim obj_AWCTeamcenterHome
'--------------------------------------------------------------------------------------------------------------------------------
'Get AWC PLM window object from xml
'--------------------------------------------------------------------------------------------------------------------------------
Set obj_AWCTeamcenterHome=Eval(GetResource("ActiveWorkspace_OR.xml").GetValue("wpage_AWCTeamcenterHome"))
'--------------------------------------------------------------------------------------------------------------------------------
'Click on user profile object
'--------------------------------------------------------------------------------------------------------------------------------
If obj_AWCTeamcenterHome.WebButton("wbtn_Your_Profile").exist Then
	 obj_AWCTeamcenterHome.WebButton("wbtn_Your_Profile").Click
End  If	 
wait 3
'--------------------------------------------------------------------------------------------------------------------------------
'Click on Sign out button
'--------------------------------------------------------------------------------------------------------------------------------
If WaitUntilExist(obj_AWCTeamcenterHome.WebButton("wbtn_SignOut"), 2, 6)Then
	If Fn_Web_UI_WebElement_Operations("Signout from Active Workspace application","Click",obj_AWCTeamcenterHome.WebButton("wbtn_SignOut"),"","","","")=True Then
		Reporter.ReportEvent micPass, "Click on [ Sign out ] button in Active Workspace", "Successfully Clicked on [ Sign out ] button in Active Workspace"
	Else
		Reporter.ReportEvent micFail, "Click on [ Sign out ] button in Active Workspace", "Fail to click on [ Sign out ] button in Active Workspace"
		ExitComponent
	End  If
End  If
wait 3
'-----------------------------------------------------------------------------------------------------------------------------------------------------------	
'Check the existance of Active Workspace page
'-----------------------------------------------------------------------------------------------------------------------------------------------------------	
If WaitUntilExist(obj_AWCTeamcenterHome.WebEdit("wedit_UserName"), 2, 10)Then
	If obj_AWCTeamcenterHome.WebEdit("wedit_UserName").exist Then
		Reporter.ReportEvent micPass, "Existence of Active Workspace application", "Successfully Signout from Active Workspace application"
	Else
		Reporter.ReportEvent micFail, "Existence of Active Workspace application", "Fail to Signout from Active Workspace application as active workspace page is exist"
		ExitComponent
	End If
End If	
Browser("Browser").Close
'-----------------------------------------------------------------------------------------------------------------------------------------------------------	
If Err.Number<> 0 Then
	Reporter.ReportEvent micFail, "Signout from Active Workspace application", "Fail to perform [ Signout Active Workspace ]  Operation due to [ " & Err.Description & " ]"
	ExitComponent
Else
	Reporter.ReportEvent micPass, "Signout from Active Workspace application", "Successfully performed [ Signout Active Workspace ] operation "
End If
'-----------------------------------------------------------------------------------------------------------------------------------------------------------	
'Set object nothing
Set obj_AWCTeamcenterHome=Nothing
'-----------------------------------------------------------------------------------------------------------------------------------------------------------	
	
