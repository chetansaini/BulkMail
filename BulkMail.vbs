' Author - Chetan Saini
' Programmer
' Recruitment Cell
' AIIMS - New Delhi
' Made with <3

' ************************************** 
' Subroutine to send emails to candidates 
' Also calls LogStatus 
' 
' Params:
' 	candidateName: name of the candidate
' 	candidateEmailID: email id of the candidate (should be in proper format)\
'	pdfFolder: folder name inside PDFs folder.
' 	pdfName: name of the attachment file (without extension)
' 	mailSubject: subject of the mail
' 	mailBody: body of the mail
' 	addAttachment: Y/N depanding on whether to add attachment or not
' **************************************
Sub EmailSend(candidateName, candidateEmailID, pdfFolder, pdfName, rank, mailSubject, mailBody, addAttachment, includeName, includeRank)
	On Error Resume Next
	objEmail.To = candidateEmailID
	objEmail.From = mailUserName
	objEmail.Subject = mailSubject
	If includeName = "Y" Then
		objEmail.HTMLBody = Replace(Replace(mailBody, "%Rank%", rank), "%Name%", candidateName)
	Else	
		objEmail.HTMLBody = mailBody
	End If
	If includeRank = "Y" Then
		objEmail.HTMLBody = Replace(Replace(mailBody, "%Rank%", rank), "%Name%", candidateName)
	Else	
		objEmail.HTMLBody = mailBody
	End If
	objEmail.Attachments.DeleteAll
	If addAttachment = "Y" Then
		objEmail.AddAttachment(currentDir & "\PDFs\" & pdfFolder & "\" & pdfName & ".pdf")
	End If
	If Err.Number = 0 Then		
		objEmail.Send
		If Err.Number = 0 Then
			Call LogStatus(candidateName, candidateEmailID, "Yes Sent")
		Else
			Call LogStatus(candidateName, candidateEmailID, Err.Description)
			'MsgBox "Cannot send mail to candidate - " & candidateName, vbCritical, "Error - Check log file after completion of execution"
			Exit Sub
		End If
	Else
		Call LogStatus(candidateName, candidateEmailID, Err.Description)
		'MsgBox "Cannot send mail to candidate - " & candidateName, vbCritical, "Error - Check log file after completion of execution"
		Exit Sub
	End If
End Sub 


' ************************************** 
' Subroutine to create a log file that 
' will store the names of candidates 
' which have been iterated over whether 
' email sent or not
' **************************************
Sub CreateLogFile()
	On Error Resume Next
	logExcelObj.Visible = False
	logExcelObj.Workbooks.Add
	logExcelObj.ActiveWorkbook.SaveAs currentDir & "\" & fso.GetFileName(excelFile) & "-log.xlsx"
	logExcelObj.ActiveWorkBook.Close
	If Err.Number <> 0 Then
		MsgBox "Cannot create log file: " & Err.Number
		excelWorkbook.close
		WScript.Quit
	End If		
End Sub


' ************************************** 
' Subroutine to configure email server 
'
' ******PLEASE DO NOT CHANGE THIS******
'
' **************************************
Sub ConfigureMailService()
	On Error Resume Next
	Const cdoSendUsingPort = 2  ' Send the message using SMTP
	Const cdoBasicAuth = 1      ' Clear-text authentication
	Const cdoTimeout = 60       ' Timeout for SMTP in seconds

	mailServer = "smtp.gmail.com"
	SMTPport = 465
	Set objEmail = CreateObject("CDO.Message")
	Set objConf = objEmail.Configuration
	Set objFlds = objConf.Fields

	With objFlds
	.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPort
	.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = mailServer
	.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = SMTPport
	.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
	.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = cdoTimeout
	.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasicAuth
	.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = mailUserName
	.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = mailPassword
	.Update
	End With
	objEmail.From = mailUserName
	If Err.Number <> 0 Then
		MsgBox "Cannot configure mail server: " & Err.Number
		excelWorkbook.close
		WScript.Quit
	End If
End Sub


' ************************************** 
' Subroutine to log candidates who have 
' been iterated over.
' Note:- It also logs reason if email is
' 	not sent.
' 
' Params:
'	candidateName: name of the candidate
'	candidateEmailID: emailId of the candidate
'	status: Yes Sent if mail is sent.
'			REASON if mail is not sent
'
' **************************************
Sub LogStatus(candidateName, candidateEmailID, status)
	On Error Resume Next
	Dim lastEntryRow
	Set logWorkbook = logExcelObj.Workbooks.open(currentDir & "\" & fso.GetFileName(excelFile) & "-log.xlsx")
	logWorkbook.Worksheets("Sheet1").Range("A"&logEntry).Value = candidateName
	logWorkbook.Worksheets("Sheet1").Range("B"&logEntry).Value = candidateEmailID
	logWorkbook.Worksheets("Sheet1").Range("C"&logEntry).Value = status
	logEntry = logEntry + 1
	logWorkbook.Save
	logWorkbook.Close
	If Err.Number <> 0 Then
		MsgBox "Cannot write logs to file: " & Err.Number
		excelWorkbook.close
		WScript.Quit
	End If
End Sub

'*******************Subroutine definations over*******************'

Dim currentDir, logFilePath, objEmail, logEntry, mailUserName
Dim candidateName, candidateEmailID, pdfName, mailSubject, mailBody, addAttachment
logEntry = 1
excelFile = WScript.Arguments(0)

Set fso = CreateObject("Scripting.FileSystemObject")
Set objExcel = CreateObject("Excel.Application")
Set logExcelObj = CreateObject("Excel.Application")

currentDir = fso.GetParentFolderName(excelFile)
excelPath = fso.GetAbsolutePathName(excelFile)

objExcel.Visible = False
Set excelWorkbook = objExcel.Workbooks.open(excelPath)
mailUserName = excelWorkbook.Worksheets("Credentials").Range("B1").Text
mailPassword = excelWorkbook.Worksheets("Credentials").Range("B2").Text

' calling subroutines
Call ConfigureMailService()
Call CreateLogFile()

MsgBox "A success message will popup when task is completed. A log file will be created which names to whom email has been sent.", vbCritical, "Tip"

' iterating through the candidates
For i = 6 To excelWorkbook.Worksheets("CandidateDetails").Range("A5").End(-4121).Row
	
	candidateName = excelWorkbook.Worksheets("CandidateDetails").Range("A"&i).Text
	candidateEmailID = excelWorkbook.Worksheets("CandidateDetails").Range("B"&i).Text
	pdfName = Trim(excelWorkbook.Worksheets("CandidateDetails").Range("C"&i).Text)
	pdfFolder = Trim(excelWorkbook.Worksheets("CandidateDetails").Range("D"&i).Text)
	rank = excelWorkbook.Worksheets("CandidateDetails").Range("E"&i).Text
	mailSubject = excelWorkbook.Worksheets("CandidateDetails").Range("B1").Text
	mailBody = excelWorkbook.Worksheets("CandidateDetails").Range("B2").Text
	addAttachment = excelWorkbook.Worksheets("CandidateDetails").Range("B3").Text
	includeName = excelWorkbook.Worksheets("CandidateDetails").Range("D3").Text
	includeRank = excelWorkbook.Worksheets("CandidateDetails").Range("F3").Text
	If Len(candidateEmailID) <> 0 Then
		Call EmailSend(candidateName, candidateEmailID, pdfFolder, pdfName, rank, mailSubject, mailBody, addAttachment, includeName, includeRank)
	Else
		Call LogStatus(candidateName, candidateEmailID, "Email ID not found.")
	End If
Next

MsgBox "Task completed for " & fso.GetFileName(excelFile), vbCritical, "Completed"
' closing the excel
excelWorkbook.close
