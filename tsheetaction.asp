<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
If Session("UIntr") = "" Then 
		Session("MSG") = "ERROR: Session has expired.<br> Please sign in again."
		Response.Redirect "default.asp"
	End If
MyNow = Now
If Request("action") = 1 Then
	If Request("confirm") <> 1 Then		
		Set rsTS = Server.CreateObject("ADODB.Recordset")
		sqlTS = "SELECT * FROM [request_T] WHERE appDate >= '" & Request("tmpDate") & "' AND appDate <= '" & Request("tmpDate2") & "' AND IntrID = " & Session("UIntr")
		rsTS.Open sqlTS, g_strCONN, 1, 3
		ctrI = Request("myctr")
			For i = 0 to ctrI + 1 
				tmpctr = Request("ctr" & i)
				If tmpctr <> "" Then
					strTmp = "index=" & tmpctr 
					tmpAppDate = GetAppDate(tmpctr)
					rsTS.Find(strTmp)
					If Not rsTS.EOF Then
						If Z_FixNull(Request("sunstart" & i)) <> "" And Z_FixNull(Request("sunend" & i)) <> "" Then
							date1st = Date & " " & Z_dates(Request("sunstart" & i))
							date2nd = Date & " " & Z_dates(Request("sunend" & i))
							If datediff("n", date1st, date2nd) >= 0 Then
								minTime = DateDiff("n", date1st, date2nd)
							Else
								minTime = DateDiff("n", date1st, dateadd("d", 1, date2nd))
							End If
							rsTS("totalhrs") = MakeTime(Z_CZero(minTime))
						Else
							rsTS("totalhrs") = Null
						End If
						If Z_fixnull(Request("sunq1" & i)) = "" then
							rsTS("happen") = Request("hid_sunq1" & i)
							happen = Request("hid_sunq1" & i)
						Else	
							rsTS("happen") = Request("sunq1" & i)
							happen = Request("sunq1" & i)
						End If
						If happen = 1 Then
							rsTS("noreas") = Request("noreasq" & i)
							rsTS("DSnoreas") = Z_CDate(Request("DSnoreas" & i))
						Else
							rsTS("noreas") = 0
							rsTS("DSnoreas") = Empty
						End If
						If Request("sunq1" & i) = 1 Then
							syscomstr = Z_fixNull(rsTS("syscom"))
							syscomstr = Replace(syscomstr, "<br>Appointment did not happen.", "")
							syscomstr = Replace(syscomstr, "<br>Appointment did happen.", "")
							rsTS("syscom") = syscomstr & "<br>Appointment did not happen." 
							rsTS("vermed") = 2
						ElseIf Request("sunq1" & i) = 2 Then
							syscomstr = Z_fixNull(rsTS("syscom"))
							syscomstr = Replace(syscomstr, "<br>Appointment did not happen.", "")
							syscomstr = Replace(syscomstr, "<br>Appointment did happen.", "")
							rsTS("syscom") = syscomstr & "<br>Appointment did happen." 
						ElseIf Request("sunq1" & i) = 0 Then
							
						End If
						If Z_FixNull(Request("sunstart" & i)) <> "" Then 
							rsTS("AStarttime") = tmpAppDate & " " & Z_dates(Request("sunstart" & i))
						Else
							rsTS("AStarttime") = Empty
						End If
						If Z_FixNull(Request("sunend" & i)) <> "" Then 
							rsTS("AEndtime") = tmpAppDate & " " & Z_dates(Request("sunend" & i))
						Else
							rsTS("AEndtime") = Empty
						End If
						rsTS("Toll") = Z_CZero(Request("suntoll" & i))
						rsTS.Update
						Call SaveHist(tmpctr, "[interpreter]timesheet.asp")
					End If
				End If
				rsTS.MoveFirst
			Next
		rsTS.Close
		Set rsTS = Nothing
	Else
		Set rsTS = Server.CreateObject("ADODB.Recordset")
		sqlTS = "SELECT * FROM [request_T] WHERE appDate >= '" & Request("tmpDate") & "' AND appDate <= '" & Request("tmpDate2") & "' AND IntrID = " & Session("UIntr")
		'On Error Resume Next
		rsTS.Open sqlTS, g_strCONN, 1, 3
		ctrI = Request("myctr")
			For i = 0 to ctrI + 1 
				tmpctr = Request("chkcon" & i)
				If tmpctr <> "" Then
					strTmp = "index=" & tmpctr 
					rsTS.Find(strTmp)
					If Not rsTS.EOF Then
						rsTS("confirmed") = MyNow
						rsTS.Update
					End If
				End If
				rsTS.MoveFirst
			Next
		rsTS.Close
		Set rsTS = Nothing
	End If
	
	If Request("confirm") = 1 Then 
		Session("MSG") = "Timesheet confirmed."
		Response.Redirect "tsheetnew.asp?tmpdate=" & Request("tmpDate")
	Else
		Session("MSG") = "Appointment saved."
		Response.Redirect "tsheet.asp?tmpdate=" & Request("tmpDate")
	End If
ElseIf Request("action") = 2 Then
	If Request("confirm") <> 1 Then
		Set rsTS = Server.CreateObject("ADODB.Recordset")
		sqlTS = "SELECT * FROM [request_T] WHERE Month(appDate) = " & Request("tmpMonth") & " AND Year(appDate) = " & Request("tmpYear") & " AND IntrID = " & Session("UIntr")
		'On Error Resume Next
		rsTS.Open sqlTS, g_strCONN, 1, 3
		ctrI = Request("myctr")
			For i = 0 to ctrI + 1 
				tmpctr = Request("ctr" & i)
				If tmpctr <> "" Then
					strTmp = "index=" & tmpctr 
					rsTS.Find(strTmp)
					If Not rsTS.EOF Then
						rsTS("toll") = Z_CZero(Request("suntoll" & i))
						rsTS.Update
					End If
				End If
				rsTS.MoveFirst
			Next
		rsTS.Close
		Set rsTS = Nothing
	Else
		Set rsTS = Server.CreateObject("ADODB.Recordset")
		sqlTS = "SELECT * FROM [request_T] WHERE Month(appDate) = " & Request("tmpMonth") & " AND Year(appDate) = " & Request("tmpYear") & " AND IntrID = " & Session("UIntr")
		'On Error Resume Next
		rsTS.Open sqlTS, g_strCONN, 1, 3
		ctrI = Request("myctr")
			For i = 0 to ctrI + 1 
				tmpctr = Request("chkcon" & i)
				If tmpctr <> "" Then
					strTmp = "index=" & tmpctr 
					rsTS.Find(strTmp)
					If Not rsTS.EOF Then
						rsTS("confirmedtoll") = MyNow
						rsTS.Update
					End If
				End If
				rsTS.MoveFirst
			Next
		rsTS.Close
		Set rsTS = Nothing
	End If
	Session("MSG") = "Mileage saved."
	If Request("confirm") = 1 Then Session("MSG") = "Mileage confirmed."
	Response.Redirect "mileage.asp?tmpMonth=" & Request("tmpMonth") & "&tmpYear=" & Request("tmpYear")
ElseIf Request("action") = 3 Then
	Set rsavail = Server.Createobject("ADODB.RecordSet")
	sqlavail = "SELECT Availability FROM Interpreter_T WHERE [index] = " & Session("UIntr")
	rsavail.Open sqlAvail, g_strCONN, 1, 3
	If not rsavail.EOF Then
		rsavail("Availability") = Request("txtAvail")
		rsavail.update
	End If
	rsavail.Close
	set rsavail = Nothing
	Session("MSG") = "Availability saved."
	Response.Redirect "avail.asp"
ElseIf Request("action") = 4 Then
	'delete entries
	set rsdel = Server.CreateObject("ADODB.RecordSet")
	sqlDel = "DELETE FROM avail_T WHERE intrID = " & Session("UIntr")
	rsDel.Open sqlDel, g_strCONN, 1, 3
	Set rsDel = Nothing
	'save entries
	Set rsavail = Server.Createobject("ADODB.RecordSet")
	sqlavail =  "SELECT * FROM avail_T"' WHERE intrID = " & Session("UIntr")
	rsavail.Open sqlAvail, g_strCONN, 1, 3
	y = 1
	Do Until y = 8
		x = 0
		Do Until x = 24
			If Request(y & x) <> "" Then
				rsavail.AddNew
				rsavail("intrID") = Session("UIntr")
				rsavail("avail") = Request(y & x)
				rsavail.Update
			End If
			x = x + 1
		Loop
		y = y + 1
	Loop
	rsavail.Close
	set rsavail = Nothing
	Session("MSG") = "Availability saved."
	Response.Redirect "avail2.asp"
End If
	
%>