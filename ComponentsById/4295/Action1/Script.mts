'******************************************************************************************************************************
' Name Of Business Component		:	Download Testfile - UFT
'
' Purpose							:	Download Testfile - UFT
'
' Input	Parameter					:	
'
' Output							:	True / False
'
' Remarks							:
'
' Author							:	Pooja Bondarde		  21 March 2022

'******************************************************************************************************************************
Option Explicit
Dim strPath,sTempFileType,strShortFileName,strfilepath,fso,oTDC,oResourceFactory,oRoot,oSub,iTotalItems,ItemCtr,CurItem
strPath = getTempFolder()
Parameter("str_TestfilePath") = strPath
Reporter.ReportEvent micPass, "File Path to store dummy files is  ["&strPath&"].", "File Path to store dummy files is  ["&strPath&"]."

If instr(Parameter("str_FileType"),".")>0 Then
	sTempFileType=Split(Parameter("str_FileType"),".")
	If sTempFileType(0)<>"" Then
		strShortFileName=Parameter("str_FileType")
	Else
		strShortFileName = "dummy"&Parameter("str_FileType")
	End If
Else
	strShortFileName = "dummy."&Parameter("str_FileType")
End If


Set fso = CreateObject("Scripting.FileSystemObject")
strfilepath = strPath&"\"&strShortFileName
If Not (fso.FileExists(strfilepath)) Then

	Set oTDC = QCUtil.QCConnection
	Set oResourceFactory = oTDC.QCResourceFactory
	Set oRoot = oResourceFactory.NewList("")
	Set oSub = Nothing
	iTotalItems = oRoot.Count
		For ItemCtr = 1 To iTotalItems
		CurItem = oRoot.Item(ItemCtr).Name
			If UCase(CurItem) = UCase(strShortFileName) Then
				Set oSub = oRoot.Item(ItemCtr)
				Exit for
			End If
		Next
	Set oRoot = Nothing
	Set oResourceFactory = Nothing
	Set oTDC = Nothing
	
	If Not oSub Is Nothing  Then
	oSub.DownloadResource strPath, True
	Parameter("str_TestfilePath") =strfilepath
	Reporter.ReportEvent micPass, "Download ["&strShortFileName&"] in File Path ["&strPath&"].", "Successfully downloaded ["&strShortFileName&"] in File Path ["&strPath&"]."
	End If
Else 
Reporter.ReportEvent micPass, "Verify ["&strShortFileName&"] is already exist in ["&strPath&"].", "Successfully verified ["&strShortFileName&"] is already exist in ["&strPath&"]."
Parameter("str_TestfilePath")=strfilepath
End If
