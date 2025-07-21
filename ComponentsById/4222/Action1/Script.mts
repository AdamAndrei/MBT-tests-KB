'******************************************************************************************************************************
' Name Of Business Component		:	Create a new revision via pirmary toolbar comand Save as or Revise - Revise
'
' Purpose							:	Create a new revision via pirmary toolbar comand Save as or Revise - Revise
'
' Input	Parameter					:	
'
' Output								:	True / False
'
' Remarks							:
'
' Author								:	Mohini Deshmukh			  13 Dec 2024

'******************************************************************************************************************************
Option Explicit
'-------------------------------------------------------------------------------------------------------------------------------
'Variable Declaration
'--------------------------------------------------------------------------------------------------------------------------------
Dim obj_AWCTeamcenterHome
Dim sTempValue,sStatusMesaage,sXpath
'--------------------------------------------------------------------------------------------------------------------------------
'Get AWC  window object from xml
'--------------------------------------------------------------------------------------------------------------------------------
Set obj_AWCTeamcenterHome=Eval(GetResource("ActiveWorkspace2406_OR.xml").GetValue("wpage_AWCTeamcenterHome"))
'--------------------------------------------------------------------------------------------------------------------------------
'Select Object  to revise
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_ObjectToRevise")<>"" Then
	Call Fn_AWC_Object_Navigation_Operations(obj_AWCTeamcenterHome,"select",Parameter("str_ObjectToRevise"))
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Click on Save As or Bulk Revise -> Bulk Revise toolbar button on primary tool bar
'--------------------------------------------------------------------------------------------------------------------------------
 Call Fn_WEB_UI_WebButton_Operations("Create a new revision via pirmary toolbar comand Save as or Revise - Revise", "click", obj_AWCTeamcenterHome, "wbtn_SaveAsorRevise","","","")
Call Fn_AWC_Common_Business_Primary_Toolbar_Operations(obj_AWCTeamcenterHome,gRevise,"")
Call  Fn_AWC_ReadyStatusSync(1)
'--------------------------------------------------------------------------------------------------------------------------------
'Enter the Detail of Revision
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_detail_of_revision") ="" Then
	Reporter.ReportEvent micFail, "Fail to enter the " & gDetailOfRevision , "Fail to enter the  " &  gDetailOfRevision  & "as [ " &  gDetailOfRevision  & " is empty ]"
	ExitComponent
Else
	call Fn_Web_UI_WebObject_Operations("Create a new revision via pirmary toolbar comand Save as or Revise - Revise", "settoproperty", obj_AWCTeamcenterHome.WebEdit("wedit_ObjectTextArea"),"", "acc_name",gDetailOfRevision)
	IF Fn_Web_UI_WebEdit_Operations("Create a new revision via pirmary toolbar comand Save as or Revise - Revise","Set",obj_AWCTeamcenterHome.WebEdit("wedit_ObjectTextArea"), "", Parameter("str_detail_of_revision")) Then
		Reporter.ReportEvent micPass, "Set the [ "&gDetailOfRevision&" ] field for Workflow Process Panel", "Successfully Set[  "& Parameter("str_detail_of_revision") &" ] for[ "&gDetailOfRevision&" ] field in Workflow Process Panel"
	Else
		Reporter.ReportEvent micFail, "Set the [ "&gDetailOfRevision&" ] field for Workflow Process Panel", "Fail to Set[  "& Parameter("str_detail_of_revision") &" ] for[ "&gDetailOfRevision&" ] field in Workflow Process Panel"
		ExitComponent
	End  IF 
	Parameter("str_detail_of_revision_out")=Parameter("str_detail_of_revision") 	
	Call Fn_AWC_ReadyStatusSync(2)	
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Add Project
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_project")<>"" Then
	'--------------------------------------------------------------------------------------------------------------------------------
	'Click on Add Project button 
	'--------------------------------------------------------------------------------------------------------------------------------
	If Fn_WEB_UI_WebButton_Operations("Create a new revision via pirmary toolbar comand Save as or Revise - Revise", "Click", obj_AWCTeamcenterHome, "wbtn_AddProject","","","") =False Then
		Reporter.ReportEvent micFail, "Click on [ Add Project ] button ", "Fail to click on [ Add Project ] button"
		ExitComponent
	Else
		Reporter.ReportEvent micPass, "Click on [ Add Project ] button", "Successfully clicked on [ Add Project ] button"
	End If
	Call  Fn_AWC_ReadyStatusSync(1)
	'--------------------------------------------------------------------------------------------------------------------------------
	'Select Project from list 
	'--------------------------------------------------------------------------------------------------------------------------------
	Call Fn_Web_UI_WebEdit_Operations("Create a new revision via pirmary toolbar comand Save as or Revise - Revise","Set",obj_AWCTeamcenterHome.WebEdit("wedit_Project_EditBox"), "", Parameter("str_project")) 
	Call  Fn_AWC_ReadyStatusSync(1)
	
	Call Fn_Web_UI_WebObject_Operations("Create a new revision via pirmary toolbar comand Save as or Revise - Revise", "settoproperty", obj_AWCTeamcenterHome.WebElement("wele_Project"), "2", "innertext", Parameter("str_project") )
	If Fn_Web_UI_WebElement_Operations("Create a new revision via pirmary toolbar comand Save as or Revise - Revise","Click",obj_AWCTeamcenterHome,"wele_Project",1,1,micLeftBtn) =False Then
		Reporter.ReportEvent micFail, "Select the Project from list", "Fail to Select the Project from list"
		ExitComponent
	Else
		Reporter.ReportEvent micPass, "Select the Project from list", "Successfully Selected the  " & Parameter("str_project") & "  Project"
	End If	
	Call  Fn_AWC_ReadyStatusSync(2)
	'--------------------------------------------------------------------------------------------------------------------------------
	'Click on Assign Project button 
	'--------------------------------------------------------------------------------------------------------------------------------
	If Fn_WEB_UI_WebButton_Operations("Create a new revision via pirmary toolbar comand Save as or Revise - Revise", "Click", obj_AWCTeamcenterHome, "wbtn_AssignProjects","","","") =False Then
		Reporter.ReportEvent micFail, "Click on [ Assign Project ] button ", "Fail to click on [ Assign Project ] button"
		ExitComponent
	Else
		Reporter.ReportEvent micPass, "Click on [ Assign Project ] button", "Successfully clicked on [ Assign Project ] button"
	End If
	Parameter("str_project_out")=Parameter("str_project")
	Call  Fn_AWC_ReadyStatusSync(2)
	'--------------------------------------------------------------------------------------------------------------------------------
	'Select Secure? check box
	'--------------------------------------------------------------------------------------------------------------------------------
	If Parameter("bln_Secure")<> "" Then
		sXpath=Replace(gCommonCheckBox2,"~",glblSecure)
		Call Fn_Web_UI_WebObject_Operations("Create a new revision via pirmary toolbar comand Save as or Revise - Revise", "settoproperty", obj_AWCTeamcenterHome.WebElement("welechk_CommonCheckBox"), "2", "xpath", sXpath)
		Call Fn_Web_UI_WebCheckBox_Operations("Create a new revision via pirmary toolbar comand Save as or Revise - Revise","set",obj_AWCTeamcenterHome,"welechk_CommonCheckBox",glblSecure,Parameter("bln_Secure"))
	End If
End If	
'--------------------------------------------------------------------------------------------------------------------------------
'Select Next Revision Notes Considered check box
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("bln_next_revision_notes_considered")<> "" Then
	sXpath=Replace(gCommonCheckBox2,"~",gNextRevisionNotesConsidered)
	Call Fn_Web_UI_WebObject_Operations("Create a new revision via pirmary toolbar comand Save as or Revise - Revise", "settoproperty", obj_AWCTeamcenterHome.WebElement("welechk_CommonCheckBox"), "2", "xpath", sXpath)
	Call Fn_Web_UI_WebCheckBox_Operations("Create a new revision via pirmary toolbar comand Save as or Revise - Revise","set",obj_AWCTeamcenterHome,"welechk_CommonCheckBox",gNextRevisionNotesConsidered,Parameter("bln_next_revision_notes_considered"))
	Call Fn_AWC_ReadyStatusSync(1)	
	Parameter("bln_next_revision_notes_considered_out")=Parameter("bln_next_revision_notes_considered")
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Enter the Comment
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_next_revision_notes") <> "" Then
	call Fn_Web_UI_WebObject_Operations("Create a new revision via pirmary toolbar comand Save as or Revise - Revise", "settoproperty", obj_AWCTeamcenterHome.WebEdit("wedit_ObjectTextArea"),"", "acc_name",gtabNextRevisionNotes)
	IF Fn_Web_UI_WebEdit_Operations("Create a new revision via pirmary toolbar comand Save as or Revise - Revise","Set",obj_AWCTeamcenterHome.WebEdit("wedit_ObjectTextArea"), "", Parameter("str_next_revision_notes")) Then
		Reporter.ReportEvent micPass, "Set the [ "&gtabNextRevisionNotes&" ] field for Workflow Process Panel", "Successfully Set[  "& Parameter("str_next_revision_notes") &" ] for[ "&gtabNextRevisionNotes&" ] field in Workflow Process Panel"
	Else
		Reporter.ReportEvent micFail, "Set the [ "&gtabNextRevisionNotes&" ] field for Workflow Process Panel", "Fail to Set[  "& Parameter("str_next_revision_notes") &" ] for[ "&gtabNextRevisionNotes&" ] field in Workflow Process Panel"
		ExitComponent
	End  IF     
	Call Fn_AWC_ReadyStatusSync(2)
	Parameter("str_next_revision_notes_out")=Parameter("str_next_revision_notes")
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Click on Revise Button
'--------------------------------------------------------------------------------------------------------------------------------
If Fn_WEB_UI_WebButton_Operations("Create a new revision via pirmary toolbar comand Save as or Revise - Revise", "click", obj_AWCTeamcenterHome, "wbtn_Revise","","","")Then
	Reporter.ReportEvent micPass, "Click on [ Revise ] button", "Successfully clicked on [  Revise ] button"
Else
	Reporter.ReportEvent micFail, "Click on [  Revise ] button", "Fail to Click on [  Revise ] button"
	ExitComponent
End  If
'--------------------------------------------------------------------------------------------------------------------------------
'Verify notification message
'--------------------------------------------------------------------------------------------------------------------------------
If obj_AWCTeamcenterHome.WebElement("wele_ObjectCreationNotificationMsg").WaitProperty("visible","true") Then
	sStatusMesaage= Fn_Web_UI_WebObject_Operations("Create a new revision via pirmary toolbar comand Save as or Revise - Revise", "getroproperty", obj_AWCTeamcenterHome.WebElement("wele_ObjectCreationNotificationMsg"), "2", "innertext", "")
	Reporter.ReportEvent micPass, "Verify Notification message", "Successfully verified notification message which is [ "&  sStatusMesaage  & " ]"
End  IF
Call Fn_AWC_ReadyStatusSync(1)
'--------------------------------------------------------------------------------------------------------------------------------
'Click on Version -> Revision History tab ,Select the Table view
'--------------------------------------------------------------------------------------------------------------------------------
 Call Fn_AWC_Object_Navigation_WorkArea_Panel_Operations(obj_AWCTeamcenterHome,"selectview_in_section","Table",gVersions & "~" & gRevisionHistory)	
'--------------------------------------------------------------------------------------------------------------------------------
'Verify Revised Revision
'--------------------------------------------------------------------------------------------------------------------------------
sTempValue= Fn_Web_UI_WebObject_Operations("Create an object", "getroproperty", obj_AWCTeamcenterHome.WebElement("wele_ObjectHeader_2"), "2", "outertext", "")

Call Fn_Web_UI_WebObject_Operations("Create a new revision via pirmary toolbar comand Save as or Revise - Revise", "settoproperty", obj_AWCTeamcenterHome.WebButton("wbtn_Section_WorkArea"), "", "outertext", gtabRevisionHistory)
Call Fn_AWC_ReadyStatusSync(1)
Call Fn_Web_UI_WebObject_Operations("Create a new revision via pirmary toolbar comand Save as or Revise - Revise", "settoproperty", obj_AWCTeamcenterHome.WebElement("wele_Objects_on_SecondaryPanel"), "", "innertext", sTempValue )
If Fn_Web_UI_WebObject_Operations("Create a new revision via pirmary toolbar comand Save as or Revise - Revise", "exist", obj_AWCTeamcenterHome.WebElement("wele_Objects_on_SecondaryPanel"), "", "", "") =False Then
	Reporter.ReportEvent micFail, "Verify the revised revision [ " & sTempValue & " ] from Section [ Versions -> Revision History ]", "Fail to verify the revised revision [ " & sTempValue & " ] from Section [ Versions -> Revision History ]"
	ExitComponent
Else
	Reporter.ReportEvent micPass, "Verify the revised revision [ " &sTempValue& " ] from Section [ Versions -> Revision History ]", "Successfully verified the revised revision [ " & sTempValue & " ] from Section [ Versions -> Revision History ]"
End If
Call Fn_AWC_ReadyStatusSync(1)
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'Store object node
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_ObjectType")<>"" Then
	'-----------------------------------------------------------------------------------------------------------------------------------------------------------
	If lcase(Parameter("str_ObjectType"))="manufacturing document" Then
		Parameter("str_Revised_Revision_Creo_out")=Replace(sTempValue,","& Split(sTempValue,",")(2),"")
		Parameter("str_Revised_Revision_out")=sTempValue
	End If
	'-----------------------------------------------------------------------------------------------------------------------------------------------------------
Else
	Parameter("str_Revised_Revision_out")=sTempValue
End If

'-----------------------------------------------------------------------------------------------------------------------------------------------------------
If Err.Number<> 0 Then
	Reporter.ReportEvent micFail, "Create a new revision via pirmary toolbar comand Save as or Revise - Revise", "Fail to perform [ Create a new revision via pirmary toolbar comand Save as or Revise - Revise ]  Operation due to [ " & Err.Description & " ]"
	ExitComponent
Else
	Reporter.ReportEvent micPass, "Create a new revision via pirmary toolbar comand Save as or Revise - Revise", "Successfully performed [ Create a new revision via pirmary toolbar comand Save as or Revise - Revise ] operation "
End If
'-----------------------------------------------------------------------------------------------------------------------------------------------------------	
'Set object nothing
Set obj_AWCTeamcenterHome=Nothing
'-----------------------------------------------------------------------------------------------------------------------------------------------------------

