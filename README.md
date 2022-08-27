# OutlookToExcelMacro
#code

	Const MACRO_NAME = "Export Messages to Excel (Rev 1)"

	Dim olkLst As Object, _
	    excApp As Object, _
	    excWkb As Object, _
	    excWks As Object, _
	    intMessages As Integer

	Sub ExportMessagesToExcel()
	    Dim strFilename As String, olkSto As Object, olkFld As Object
	    strFilename = InputBox("Enter a filename (including path) to save the exported messages to.", MACRO_NAME)
	    If strFilename <> "" Then
	        intMessages = 0
	        Set excApp = CreateObject("Excel.Application")
	        excApp.Visible = False
	        Set excWkb = excApp.Workbooks.Add
	        'Set excWks = excWkb.Worksheets(1)
	        'Set excWks = excWkb.Worksheets.Add()
	        Set excWks = excWkb.ActiveSheet
	        For Each olkSto In Session.Stores
	            'Write Excel Column Headers
	            With excWks
	                .Cells(1, 1) = "Folder"
	                .Cells(1, 2) = "Received"
	                .Cells(1, 3) = "Unread"
	                .Cells(1, 4) = "Sender Name"
	                .Cells(1, 5) = "Subject"
	            End With
	        Set olkSto = GetNamespace("MAPI")
	        Set olkFld = olkSto.PickFolder
	        ProcessFolder olkFld
	    Next
	        excWkb.SaveAs strFilename
	    End If
	    Set excWks = Nothing
	    Set excWkb = Nothing
	    excApp.Quit
	    Set excApp = Nothing
	    MsgBox "Process complete.  A total of " & intMessages & " messages were exported.", vbInformation + vbOKOnly, MACRO_NAME
	End Sub
	 
	Sub ProcessFolder(olkFld As Outlook.MAPIFolder)
	    Dim olkMsg As Variant, olkSub As Outlook.MAPIFolder, intRow As Integer
	    intRow = excWks.UsedRange.Rows.Count
	    intRow = intRow + 1

	    'Trick To Avoid Break runtime error 430 - "Class Does Not Support Automation or Expected Interface" Error Message 
	    On Error Resume Next

	    'Write messages to spreadsheet
	    For Each olkMsg In olkFld.Items
	        'Only export messages, not receipts or appointment requests, etc.
	        If olkMsg.Class = olMail Then
	            'Add a row for each field in the message you want to export
	            excWks.Cells(intRow, "A").Value = olkFld.Name
	            excWks.Cells(intRow, "B").Value = olkMsg.ReceivedTime
	            excWks.Cells(intRow, "C").Value = olkMsg.UnRead
	            excWks.Cells(intRow, "D").Value = olkMsg.SenderName
	            excWks.Cells(intRow, "E").Value = olkMsg.Subject

	            intRow = intRow + 1
	            intMessages = intMessages + 1
	        End If
	    Next
	    Set olkMsg = Nothing
	    For Each olkSub In olkFld.Folders
	        ProcessFolder olkSub
	    Next
	    Set olkLst = Nothing
	    Set olkSub = Nothing
	End Sub


# runtime error 430 in this macro happens when found an unexpected empty Row like recalled email.
# This Outlook Macro - Export Messages to Excel By Inserting File Path and name and by determine Which inbox and its sub to export.
https://docs.microsoft.com/en-us/office/vba/api/outlook.mailitem#properties

https://www.linkedin.com/embed/feed/update/urn:li:ugcPost:6968899634085707776?compact=1

