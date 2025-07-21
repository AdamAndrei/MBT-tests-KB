'******************************************************************************************************************************
' Name Of Business Component		:	 Create a TR Document By Tile Create 
'
' Purpose							:	Perform the Create a TR Document By Tile Create  operation
'
' Input	Parameter					:	
'
' Output							:	True / False
'
' Remarks							:
'
' Author							:	Mohini Deshmukh 			  22 March 2022

'******************************************************************************************************************************
Option Explicit
'--------------------------------------------------------------------------------------------------------------------------------
'Variable Declaration
'--------------------------------------------------------------------------------------------------------------------------------
Dim objDic,sOjectNodeName
Set objDic=CreateObject("Scripting.dictionary")
'--------------------------------------------------------------------------------------------------------------------------------
'Enter Title  value
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_Title") ="" Then
	Reporter.ReportEvent micFail, "Enter the " & gTitle & " value", "Create Business object of type  " & sObjType & " as [ " & gTitle & " value is Empty ] "
	ExitComponent
End If
objDic.Add gTitle,Parameter("str_Title")
Parameter("str_Title_out") =Parameter("str_Title") 
'--------------------------------------------------------------------------------------------------------------------------------
'Select Language value
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_Language") ="" Then
	Reporter.ReportEvent micFail, "Enter the " & gLanguage & " value", "Create Business object of type  " & sObjType & " as [ " & gLanguage & " value is Empty ] "
	ExitComponent
End If
objDic.Add gLanguage2,Parameter("str_Language")
Parameter("str_Language_out") =Parameter("str_Language")
'--------------------------------------------------------------------------------------------------------------------------------
'Select Document Group
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_Document_Group") ="" Then
	Reporter.ReportEvent micFail, "Enter the " & gDocumentGroup & " value", "Create Business object of type  " & sObjType & " as [ " & gDocumentGroup & " value is Empty ] "
	ExitComponent
End If
objDic.Add gDocumentGroup,Parameter("str_Document_Group")
Parameter("str_Document_Group_out") = Parameter("str_Document_Group")
'--------------------------------------------------------------------------------------------------------------------------------
'Select Document Kind
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_Document_Kind") ="" Then
	Reporter.ReportEvent micFail, "Enter the " & gDocumentKind & " value", "Create Business object of type  " & sObjType & " as [ " & gDocumentKind & " value is Empty ] "
	ExitComponent
End If
objDic.Add gDocumentKind,Parameter("str_Document_Kind")
 Parameter("str_Document_Kind_out") =Parameter("str_Document_Kind")
 '--------------------------------------------------------------------------------------------------------------------------------
'Select Test Category
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_Test_Category") ="" Then
	Reporter.ReportEvent micFail, "Enter the  " & gTestCategory & " Value", "Create Business object of type  " & sObjType & " as [  " & gTestCategory & " Value is Empty ] "
	ExitComponent
End If
objDic.Add gTestCategory,Parameter("str_Test_Category")
 Parameter("str_Test_Category_out") =Parameter("str_Test_Category")
 '--------------------------------------------------------------------------------------------------------------------------------
'Select Billing Type
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_Billing_Type") ="" Then
	Reporter.ReportEvent micFail, "Enter the " & gBillingType & " Value", "Create Business object of type  " & sObjType & " as [ " & gBillingType & " Value is Empty ] "
	ExitComponent
End If
objDic.Add gBillingType,Parameter("str_Billing_Type")
 Parameter("str_Billing_Type_out") =Parameter("str_Billing_Type")
 '--------------------------------------------------------------------------------------------------------------------------------
'Select Test Status
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_Test_Status")<>"" Then
objDic.Add gTestStatus,Parameter("str_Test_Status")
 Parameter("str_Test_Status_out") =Parameter("str_Test_Status")
 End If
 '--------------------------------------------------------------------------------------------------------------------------------
'Select Requested By
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_Requested_By")<>"" Then
objDic.Add gRequestedBy,Parameter("str_Requested_By")
 Parameter("str_Requested_By_out") =Parameter("str_Requested_By")
 End If
 '--------------------------------------------------------------------------------------------------------------------------------
'Select Responsible CoC
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_Responsible_CoC") ="" Then
	Reporter.ReportEvent micFail, "Enter the " & gResponsibleCoC & " Value", "Create Business object of type  " & sObjType & " as [ " & gResponsibleCoC & " Value is Empty ] "
	ExitComponent
End If
objDic.Add gResponsibleCoC,Parameter("str_Responsible_CoC")
 Parameter("str_Responsible_CoC_out") =Parameter("str_Responsible_CoC")
 '--------------------------------------------------------------------------------------------------------------------------------
'Select Requested by Location
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_Requested_by_Location")<>"" Then
objDic.Add gRequestedByLocation,Parameter("str_Requested_by_Location")
 Parameter("str_Requested_by_Location_out") =Parameter("str_Requested_by_Location")
 End If
'--------------------------------------------------------------------------------------------------------------------------------
'Select Department Code
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_Department_Code") <>"" Then
objDic.Add gDepartmentCode,Parameter("str_Department_Code")
 Parameter("str_Department_Code_out") = Parameter("str_Department_Code")
 End If
'--------------------------------------------------------------------------------------------------------------------------------
'Enter Estimated Hours
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_Estimated_Hours") <>"" Then
objDic.Add gEstimatedHours,Parameter("str_Estimated_Hours")
 Parameter("str_Estimated_Hours_out") = Parameter("str_Estimated_Hours")
 End If
 '--------------------------------------------------------------------------------------------------------------------------------
'Enter Date Needed value
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_Date_Needed") <>"" Then
objDic.Add gDateNeeded,Parameter("str_Date_Needed")
 Parameter("str_Date_Needed_out") = Parameter("str_Date_Needed")
 End If
 '--------------------------------------------------------------------------------------------------------------------------------
'Enter Objective value
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_Objective") ="" Then
	Reporter.ReportEvent micFail, "Enter the " & gObjective & " Value", "Create Business object of type  " & sObjType & " as [ " & gObjective & " Value is Empty ] "
	ExitComponent
End If
objDic.Add gObjective,Parameter("str_Objective")
 Parameter("str_Objective_out") =Parameter("str_Objective")
 '--------------------------------------------------------------------------------------------------------------------------------
'Enter Detail of Revision value
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_Detail_of_Revision") <>"" Then
objDic.Add gDetailOfRevision,Parameter("str_Detail_of_Revision")
 Parameter("str_Detail_of_Revision_out") = Parameter("str_Detail_of_Revision")
 End If
 '--------------------------------------------------------------------------------------------------------------------------------
'Select Team
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_Team") ="" Then
	Reporter.ReportEvent micFail, "Enter the " & gTeam & " Value", "Create Business object of type  " & sObjType & " as [ " & gTeam & " Value is Empty ] "
	ExitComponent
End If
objDic.Add gTeam,Parameter("str_Team")
 Parameter("str_Team_out") = Parameter("str_Team")
' ' '--------------------------------------------------------------------------------------------------------------------------------
' 'Organization field is not availble anymore - 17 Jun 2024 - Modified By Mohini
' '--------------------------------------------------------------------------------------------------------------------------------
''Select Organizations
''--------------------------------------------------------------------------------------------------------------------------------
'If Parameter("str_Organizations") <>"" Then
'	objDic.Add gOrganizations,Parameter("str_Organizations")
'	 Parameter("str_Organizations_out") = Parameter("str_Organizations")
'End If
'--------------------------------------------------------------------------------------------------------------------------------
'Select Locations
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_Locations") <>"" Then
objDic.Add gLocations,Parameter("str_Locations")
 Parameter("str_Locations_out") = Parameter("str_Locations")
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Select All Location Access check box
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("bln_all_location_access")<>"" Then
	objDic.Add gAllLocationAccess,Parameter("bln_all_location_access")
	Parameter("bln_All_Location_Access_out")=Parameter("bln_all_location_access")
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Select Engineer Responsible
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_EngineerResponsible")<>"" Then
	objDic.Add gEngineerResponsible,Parameter("str_EngineerResponsible")
	Parameter("str_EngineerResponsible_out")=Parameter("str_EngineerResponsible")
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Select Manager Responsible
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_ManagerResponsible")<>"" Then
	objDic.Add gManagerResponsible,Parameter("str_ManagerResponsible")
	Parameter("str_ManagerResponsible_out")=Parameter("str_ManagerResponsible")
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Select  Responsible Tester
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_ResponsibleTester")<>"" Then
	objDic.Add gResponsibleTester,Parameter("str_ResponsibleTester")
	Parameter("str_ResponsibleTester_out")=Parameter("str_ResponsibleTester")
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Enter Engineering Project Number value
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_Engineering_Project_Number") <>"" Then
	objDic.Add gEngineeringProjectNumber,Parameter("str_Engineering_Project_Number")
	Parameter("str_Engineering_Project_Number_out")=Parameter("str_Engineering_Project_Number")
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Enter Engineering Project due Date value
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_Engineering_Project_due_Date") <>"" Then
	objDic.Add gEngineeringProjectDueDate,Parameter("str_Engineering_Project_due_Date")
	Parameter("str_Engineering_Project_due_Date_out")=Parameter("str_Engineering_Project_due_Date")
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Enter Testing Completion Date value
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_Testing_Completion_Date") <>"" Then
	objDic.Add gTestingCompletionDate,Parameter("str_Testing_Completion_Date")
	Parameter("str_Testing_Completion_Date_out")=Parameter("str_Testing_Completion_Date")
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Select Customer
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_Customer")<>"" Then
	objDic.Add gCustomer,Parameter("str_Customer")
	Parameter("str_Customer_out")=Parameter("str_Customer")
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Select Device Code
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_Device_Code")<>"" Then
	objDic.Add gDeviceCode,Parameter("str_Device_Code")
	Parameter("str_Device_Code_out")=Parameter("str_Device_Code")
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Select Test Location
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_Test_Location")<>"" Then
	objDic.Add gTestLocation,Parameter("str_Test_Location")
	Parameter("str_Test_Location_out")=Parameter("str_Test_Location")
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Select Test Part Status
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_Test_Part_Status")<>"" Then
	objDic.Add gTestPartStatus,Parameter("str_Test_Part_Status")
	Parameter("str_Test_Part_Status_out")=Parameter("str_Test_Part_Status")
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Select Test Part Disposition
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_Test_Part_Disposition")<>"" Then
	objDic.Add gTestPartDisposition,Parameter("str_Test_Part_Disposition")
	Parameter("str_Test_Part_Disposition_out")=Parameter("str_Test_Part_Disposition")
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Enter Number Of Samples value
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_Number_Of_Samples") <>"" Then
	objDic.Add gNumberOfSamples,Parameter("str_Number_Of_Samples")
	Parameter("str_Number_Of_Samples_out")=Parameter("str_Number_Of_Samples")
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Enter Vendor value
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_Vendor") <>"" Then
	objDic.Add gVendor,Parameter("str_Vendor")
	Parameter("str_Vendor_out")=Parameter("str_Vendor")
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Enter Charge Number value
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_Charge_Number") <>"" Then
	objDic.Add gChargeNumber,Parameter("str_Charge_Number")
	Parameter("str_Charge_Number_out")=Parameter("str_Charge_Number")
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Enter Comments value
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_Comments") <>"" Then
	objDic.Add gComments,Parameter("str_Comments")
	Parameter("str_Comments_out") = Parameter("str_Comments")
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Enter Non TC Specification value
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_Non_TC_Specification") <>"" Then
	objDic.Add gNonTCSpecification,Parameter("str_Non_TC_Specification")
	Parameter("str_Non_TC_Specification_out") = Parameter("str_Non_TC_Specification")
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Enter Vehicle Description value
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_Vehicle_Description") <>"" Then
	objDic.Add gVehicleDescription,Parameter("str_Vehicle_Description")
	Parameter("str_Vehicle_Description_out") = Parameter("str_Vehicle_Description")
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Enter Vehicle Make value
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_Vehicle_Make") <>"" Then
	objDic.Add gVehicleMake,Parameter("str_Vehicle_Make")
	Parameter("str_Vehicle_Make_out") = Parameter("str_Vehicle_Make")
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Enter Vehicle Model value
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_Vehicle_Model") <>"" Then
	objDic.Add gVehicleModel,Parameter("str_Vehicle_Model")
	Parameter("str_Vehicle_Model_out") = Parameter("str_Vehicle_Model")
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Enter VIN Number value
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_VIN_Number") <>"" Then
	objDic.Add gVINNumber,Parameter("str_VIN_Number")
	Parameter("str_VIN_Number_out") = Parameter("str_VIN_Number")
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Enter ABS Software Revision  value
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_ABS_Software_Revision") <>"" Then
	objDic.Add gABSSoftwareRevision,Parameter("str_ABS_Software_Revision")
	Parameter("str_ABS_Software_Revision_out") = Parameter("str_ABS_Software_Revision")
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Enter ESP Software Revision  value
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_ESP_Software_Revision") <>"" Then
	objDic.Add gESPSoftwareRevision,Parameter("str_ESP_Software_Revision")
	Parameter("str_ESP_Software_Revision_out") = Parameter("str_ESP_Software_Revision")
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Enter DAS Software Revision  value
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_DAS_Software_Revision") <>"" Then
	objDic.Add gDASSoftwareRevision,Parameter("str_DAS_Software_Revision")
	Parameter("str_DAS_Software_Revision_out") = Parameter("str_DAS_Software_Revision")
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Select Component Code
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_Component_Code")<>"" Then
	objDic.Add gComponentCode,Parameter("str_Component_Code")
	Parameter("str_Component_Code_out") = Parameter("str_Component_Code")
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Select Test Classification
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_Test_Classification")<>"" Then
	objDic.Add gTestClassification,Parameter("str_Test_Classification")
	Parameter("str_Test_Classification_out") = Parameter("str_Test_Classification")
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Enter Background  value
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_Background") <>"" Then
	objDic.Add gBackground,Parameter("str_Background")
	Parameter("str_Background_out") = Parameter("str_Background")
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Enter Test Comments and Conclusions  value
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_Test_Comments_and_Conclusions") <>"" Then
	objDic.Add gTestCommentsAndConclusions,Parameter("str_Test_Comments_and_Conclusions")
	Parameter("str_Test_Comments_and_Conclusions_out") = Parameter("str_Test_Comments_and_Conclusions")
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Enter Corrective Action  value
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_Corrective_Action") <>"" Then
	objDic.Add gCorrectiveAction,Parameter("str_Corrective_Action")
	Parameter("str_Corrective_Action_out") = Parameter("str_Corrective_Action")
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Enter Material Comments  value
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_Material_Comments") <>"" Then
	objDic.Add gMaterialComments,Parameter("str_Material_Comments")
	Parameter("str_Material_Comments_out") = Parameter("str_Material_Comments")
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Enter Procedure Comments  value
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_Procedure_Comments") <>"" Then
	objDic.Add gProcedureComments,Parameter("str_Procedure_Comments")
	Parameter("str_Procedure_Comments_out") = Parameter("str_Procedure_Comments")
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Enter Fixture Comments  value
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_Fixture_Comments") <>"" Then
	objDic.Add gFixtureComments,Parameter("str_Fixture_Comments")
	Parameter("str_Fixture_Comments_out") = Parameter("str_Fixture_Comments")
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Select Material check box
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("bln_Material")<>"" Then
	objDic.Add gMaterial2,Parameter("bln_Material")
	Parameter("bln_Material_out") = Parameter("bln_Material")
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Select Procedure check box
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("bln_Procedure")=True Then
	objDic.Add gProcedure,Parameter("bln_Procedure")
	Parameter("bln_Procedure_out") = Parameter("bln_Procedure")
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Select Fixture check box
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("bln_Fixture")=True Then
	objDic.Add gFixture,Parameter("bln_Fixture")
	Parameter("bln_Fixture_out") = Parameter("bln_Fixture")
End If
'--------------------------------------------------------------------------------------------------------------------------------
'Add Project
'--------------------------------------------------------------------------------------------------------------------------------
If Parameter("str_assgin_projects") <>"" Then
	objDic.Add gProject,Parameter("str_assgin_projects")
	Parameter("str_Assgin_Projects_out") =Parameter("str_assgin_projects")
End If	
'--------------------------------------------------------------------------------------------------------------------------------
'Store the search result
'--------------------------------------------------------------------------------------------------------------------------------
sOjectNodeName= Split(Fn_AWC_Object_Creation_Operation("bytile",Parameter("str_type_of_object"),objDic,Parameter("str_Performance_Monitor") ),"~")
Parameter("str_newly_created_object_out")=sOjectNodeName(0)
Parameter("str_ID_out")=Split(sOjectNodeName(0),",")(0)
objDic.RemoveAll
'---------------------------------------------------------------------------------------------------------------------------------
