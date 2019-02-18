<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<!-- #include file="_Security.asp" -->
<%

If Z_GetInfoFROMAppID(Request("ReqID"), "IntrID") = Session("UIntr") Then
	'create vform
	Set theDoc = Server.CreateObject("ABCpdf9.Doc") 'converts html to pdf
	Set theDoc2 = Server.CreateObject("ABCpdf9.Doc")
	Set theDoc3 = Server.CreateObject("ABCpdf9.Doc")
	
	fname = "VerificationForm" & Request("ReqID") & ".pdf"
	attachPDF = pdfStr & fname
	strUrl = "https://interpreter.thelanguagebank.org/interpreter/print.asp?PDF=1&ID=" & Request("ReqID") & "&IDIntr=" & Session("UIntr")
	thedoc.HtmlOptions.PageCacheClear
	theDoc.HtmlOptions.RetryCount = 3
	theDoc.HtmlOptions.Timeout = 120000
	theDoc.Pos.X = 10
	theDoc.Pos.Y = 10
	theID = theDoc.AddImageUrl(strUrl)
	
	Do
	  If Not theDoc.Chainable(theID) Then Exit Do
	  theDoc.Page = theDoc.AddPage()
	  theID = theDoc.AddImageToChain(theID)
	Loop
	If Z_GetInfoFROMAppID(Request("ReqID"), "InstID") = 671 Then 'saint vincent
		theDoc.FontSize = 12 ' big text
		theDoc.rect.Move 50, -50
		theDoc.Page = theDoc.AddPage(1)
		theText = "<b>ATTENTION INTERPRETERS</b><br><br><br><br>" & _
			"When handling appointments at St. Vincent’s Hospital, you must follow the below procedures:<br><br><br>" & _ 
			"1)	BEFORE THE APPOINTMENT: Go to the mailroom. The mailroom is by the Loading Dock/Receiving Area.<br><br>" & _ 
			"2) Ask for Ms. Fran Goulet (phone number 508-363-9310 if you need to contact her.)<br><br>" & _ 
			"3) Sign the log book, take the badge Ms. Goulet hands you, and continue on to the appointment.<br><br>" & _ 
			"4) After the appointment, make sure that all parts of the V-Form are correctly filled out. If any part of the<br>" & _
			"V-Form is incomplete, we cannot bill St. Vincent’s!<br><br>" & _ 
			"5)	Once the V-Form is complete, leave it with the Interpretation Department on the ground floor. DO NOT ASK<br>" & _
			"FOR A COPY. We do not need copies of V-Forms from St. Vincent’s.<br><br><br>" & _ 
			"THANK YOU VERY MUCH FOR FOLLOWING THIS PROCEDURE."
		theDoc.AddHtml(theText)
	ElseIf Z_GetInfoFROMAppID(Request("ReqID"), "InstID") = 849 Then 'lowel
		theDoc.FontSize = 12 ' big text
		theDoc.rect.Move 50, -50
		theDoc.Page = theDoc.AddPage(1)
		theText = "Instructions for assignments at Lowell General Hospital:<br><br><br><br>" & _
			"1.	Assignments are scheduled for 2 hours.  You must be available to stay for the full 2-hours, since new interpreting<br>" & _
			"sessions can be assigned to us little or no notice, and we may need you to complete them.<br><br>" & _
			"2.	Upon completion of an interpreting assignment, please contact Interpreter Services by dialing extension 64710 or<br>" & _
			"64709 (Saints Campus) or extension 76591 (Main Campus) for further instructions.<br><br>" & _
			"3.	If the duration of the assignment is expected to exceed 2-hours, please call the Interpreter Services office at<br>" & _
			"extension 64710 or 64709 (Saints) or extension 76591 (Main) for approval to stay longer.<br><br>" & _
			"4.	If an appointment is cancelled or the patient does not show up, please dial extension 64710 or 64709 (Saints) or<br>" & _
			"extension 76591 (Main) for further instructions. We may need you for another appointment elsewhere in the hospital.<br><br>" & _
			"5.	Upon completion of an appointment at one of our satellite clinics, please call the Interpreter Services office at<br>" & _
			"extension 64710 or 64709 (Saints) or extension 76591 (Main) to provide us with information, especially when the<br>" & _
			"appointment did not go as planned, the patient didn’t show, it started late, etc.<br><br>" & _
			"&nbsp;&nbsp;&nbsp;&nbsp;•	If nobody is available to take your call when you contact the Interpreter Services office, please leave<br>" & _
			"a message with details about the appointment.<br><br>" & _
			"Main Office 978-937-6591"
		theDoc.AddHtml(theText)
	ElseIf Z_GetInfoFROMAppID(Request("ReqID"), "InstID") = 108 Then 'dhhs
		LangID = Z_GetInfoFROMAppID(Request("ReqID"), "LangID")
		theDoc2.Read(SurveyPath & GetLangSurvey(LangID))
		theDoc.Append(theDoc2) 
	ElseIf Z_GetInfoFROMAppID(Request("ReqID"), "InstID") = 323 And DeptID = 1924 Then 'wentworth
		theDoc2.Read(DirectionPath & "DirWDH-CNS.pdf")
		theDoc.Append(theDoc2) 
	ElseIf Z_GetInfoFROMAppID(Request("ReqID"), "InstID") = 860 Then 'umass
		theDoc2.Read(DirectionPath & "READ ME FIRST.pdf")
		theDoc.Append(theDoc2)
		theDoc3.Read(DirectionPath & "Interpreters guidelines.pdf")
		theDoc.Append(theDoc3)
	End If
	For i = 1 To theDoc.PageCount
	  theDoc.PageNumber = i
	  theDoc.Flatten
	Next
	
	theDoc.Save attachPDF
	
	Set theDoc3 = Nothing
	Set theDoc2 = Nothing
	Set theDoc = Nothing 
	
	'save datestamp adn save in hist
	Set rsTS = Server.CreateObject("ADODB.RecordSet")
	rsTS.Open "UPDATE request_T SET vformdownload = '" & Now & "' WHERE [index] = " & Request("ReqID"), g_strCONN, 1, 3
	Set rsTS = Nothing
	'response.write "ID: " & Request("ReqID")
	Call SaveHist(Request("ReqID"), "[interpreter]dl.asp") 
	
	
	'downloadfile
	'Set dload = Server.CreateObject("SCUpload.Upload")
	'tmpfile = attachPDF ' "C:\work\LSS-LBIS\web-DMZ\Images\icon_download.gif" 'attachPDF
	'dload.Download tmpFile
	'Set dload = Nothing
	download = 1
	
	tmpfile = attachPDF
	Set objStream = Server.CreateObject("ADODB.Stream")
  objStream.Type = 1 'adTypeBinary
  objStream.Open
  objStream.LoadFromFile(tmpfile)
  Response.ContentType = "application/x-unknown"
  Response.Addheader "Content-Disposition", "attachment; filename=" & fname 
  Response.BinaryWrite objStream.Read
  objStream.SaveToFile tmpFile, 2
  objStream.Close
  Set objStream = Nothing
Else
	response.write "There was an error in creating the Verification Form. Please close this browser and try again try again later."
End If

%>