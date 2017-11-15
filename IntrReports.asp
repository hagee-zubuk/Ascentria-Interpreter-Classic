<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<!-- #include file="_Security.asp" -->
<%
Function GetTrain(xxx)
	GetTrain = ""
	Set rsTrain = Server.CreateOBject("ADODB.RecordSet")
	sqlTrain = "SELECT * FROM  Training_T WHERE [index] = " & xxx
	rsTrain.Open sqlTrain, g_strCONN, 1, 3
	If Not rsTrain.EOF Then
		GetTrain = rsTrain("training")
	End If
	rsTrain.Close
	Set rsTrain = Nothing
End Function
server.scripttimeout = 360000
tmpDate2 = date
tmpReport = Split(Z_DoDecrypt(Request.Cookies("LBREPORT")), "|")
tmpdate = replace(date, "/", "") 
tmpTime = replace(FormatDateTime(time, 3), ":", "")
If Request("ctrl") = 1 Then
	RepCSV =  "ExpireDocs" & tmpdate & ".csv" 
	strMSG = "Expiring Interpreter Documents report."
	strHead = "<td class='tblgrn'>Name</td>" & vbCrlf & _
		"<td class='tblgrn'>Document</td>" & vbCrlf & _
		"<td class='tblgrn'>Expiration Date</td>" & vbCrlf 
	CSVHead = "Last Name, First Name, Document, Expiration Date"
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	sqlRep = "SELECT * FROM Interpreter_T WHERE Active = 1 ORDER BY [last name], [first name]"
	rsRep.Open sqlRep, g_strCONN,1 ,3
	y = 0
	Do Until rsRep.EOF
		tmpName = rsRep("last name") & ", " & rsRep("first name")
		kulay = "#FFFFFF"
		If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
		If Not IsNull(rsRep("passexp")) Then 
			If DateDiff("d", tmpDate2, rsRep("passexp")) < 15 Then 
				strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & tmpName & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>Passport</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & rsRep("passexp") & "</td></tr>" & vbCrLf
				CSVBody = CSVBody & rsRep("last name") & "," & rsRep("first name") & ",Passport," & rsRep("passexp") & vbCrLf
				y = y + 1
			End If
		End If
		If Not IsNull(rsRep("driveexp")) Then 
			If DateDiff("d", tmpDate2, rsRep("driveexp")) < 15 Then 
				strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & tmpName & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>Driver's License</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & rsRep("driveexp") & "</td></tr>" & vbCrLf
				CSVBody = CSVBody & rsRep("last name") & "," & rsRep("first name") & ",Driver's License," & rsRep("driveexp") & vbCrLf
				y = y + 1
			End If
		End If
		If Not IsNull(rsRep("greenexp")) Then 
			If DateDiff("d", tmpDate2, rsRep("greenexp")) < 15 Then 
				strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & tmpName & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>Green Card</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & rsRep("greenexp") & "</td></tr>" & vbCrLf
				CSVBody = CSVBody & rsRep("last name") & "," & rsRep("first name") & ",Green Card," & rsRep("greenexp") & vbCrLf
				y = y + 1
			End If
		End If
		If Not IsNull(rsRep("employexp")) Then 
			If DateDiff("d", tmpDate2, rsRep("employexp")) < 15 Then 
				strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & tmpName & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>Employment Authorization</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & rsRep("employexp") & "</td></tr>" & vbCrLf
				CSVBody = CSVBody & rsRep("last name") & "," & rsRep("first name") & ",Employment Authorization," & rsRep("employexp") & vbCrLf
				y = y + 1
			End If
		End If
		If Not IsNull(rsRep("carexp")) Then 
			If DateDiff("d", tmpDate2, rsRep("carexp")) < 15 Then 
				strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & tmpName & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>Car Insurance</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & rsRep("carexp") & "</td></tr>" & vbCrLf
				CSVBody = CSVBody & rsRep("last name") & "," & rsRep("first name") & ",Car Insurance," & rsRep("carexp") & vbCrLf
				y = y + 1
			End If
		End If
		rsRep.MoveNext
	Loop
	rsRep.Close
	Set rsRep = Nothing
ElseIf Request("ctrl")= 2 Then
	If Request("selRep") = 1 Then 'training
		RepCSV =  "IntrTrain" & tmpdate & ".csv" 
		strMSG = "Interpreter Training"
		If Request("txtyear") <> 0 Then
			strMSG = strMSG & " for the year " & Request("txtyear")
			tmpDate1 = cdate("1/1/" & Request("txtyear"))
		End IF
		strHead = "<td class='tblgrn'>Name</td>" & vbCrlf & _
			"<td class='tblgrn'>Date</td>" & vbCrlf & _
			"<td class='tblgrn'>Hours</td>" & vbCrlf & _
			"<td class='tblgrn'>Training</td>" & vbCrlf 
		CSVHead = "Last Name, First Name, Date, Hours, Training"
		If IsDate(tmpDate2) Then
			tmpYear = Year(tmpDate1)
			Set rsRep = Server.CreateObject("ADODB.RecordSet")
			If Request("txtyear") <> 0 Then
				sqlRep = "SELECT * FROM IntrTraining_T, interpreter_T WHERE Year(date) = " & tmpYear & _
					" AND intrID = interpreter_T.[index] AND Active = 1 ORDER BY [last name], [first name], date"
			Else
				sqlRep = "SELECT * FROM IntrTraining_T, interpreter_T WHERE Active = 1 AND intrID = interpreter_T.[index] ORDER BY [last name], [first name], date"
			End If
			rsRep.Open sqlRep, g_strCONN, 1, 3
			Do Until rsRep.EOF
				tmpName = rsRep("last name") & ", " & rsRep("first name")
				kulay = "#FFFFFF"
				If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
				strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & tmpName & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & rsRep("date") & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & rsRep("Hours") & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & GetTrain(rsRep("Type")) & "</td></tr>" & vbCrLf
				CSVBody = CSVBody & rsRep("last name") & "," & rsRep("first name") & "," & rsRep("date") & "," & rsRep("Hours") & "," & GetTrain(rsRep("Type")) & vbCrLf
				rsRep.MoveNext
			Loop
			rsRep.Close
			Set rsRep = Nothing 
		End If
	ElseIf Request("selRep") = 2 Then'eval/feed 
		RepCSV =  "IntrEval" & tmpdate & ".csv" 
		strMSG = "Interpreter Evaluation/Feedback"
		If Request("txtyear") <> 0 Then
			strMSG = strMSG & " for the year " & Request("txtyear")
			tmpDate1 = cdate("1/1/" & Request("txtyear"))
		End IF
		strHead = "<td class='tblgrn'>Name</td>" & vbCrlf & _
			"<td class='tblgrn'>Date</td>" & vbCrlf & _
			"<td class='tblgrn'>Training</td>" & vbCrlf 
		CSVHead = "Last Name, First Name, Date, Evaluation/Feedback"
		If IsDate(tmpDate2) Then
			tmpYear = Year(tmpDate1)
			Set rsRep = Server.CreateObject("ADODB.RecordSet")
			If Request("txtyear") <> 0 Then
				sqlRep = "SELECT * FROM InterpreterEval_T, interpreter_T WHERE Year(date) = " & tmpYear & _
					" AND intrID = interpreter_T.[index] AND Active = 1 ORDER BY [last name], [first name], date"
			Else
				sqlRep = "SELECT * FROM InterpreterEval_T, interpreter_T WHERE Active = 1 AND intrID = interpreter_T.[index] ORDER BY [last name], [first name], date"
			End If
			rsRep.Open sqlRep, g_strCONN, 1, 3
			Do Until rsRep.EOF
				tmpName = rsRep("last name") & ", " & rsRep("first name")
				kulay = "#FFFFFF"
				If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
				strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & tmpName & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & rsRep("date") & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & rsRep("comment") & "</td></tr>" & vbCrLf
				CSVBody = CSVBody & rsRep("last name") & "," & rsRep("first name") & "," & rsRep("date") & "," & rsRep("comment") & vbCrLf
				y = y + 1
				rsRep.MoveNext
			Loop
			rsRep.Close
			Set rsRep = Nothing 
		End If
	ElseIf Request("selRep") = 3 Then'docs
		RepCSV =  "IntrDocs" & tmpdate & ".csv" 
		strMSG = "Interpreter Documents" 
		strHead = "<td class='tblgrn'>Name</td>" & vbCrlf & _
			"<td class='tblgrn'>Document</td>" & vbCrlf & _
			"<td class='tblgrn'>Number</td>" & vbCrlf & _
			"<td class='tblgrn'>Expiration Date</td>" & vbCrlf 
		CSVHead = "Last Name, First Name, Document, Number, Expiration Date"
		Set rsRep = Server.CreateObject("ADODB.RecordSet")
		sqlRep = "SELECT * FROM Interpreter_T WHERE Active = 1 ORDER BY [last name], [first name]"
		rsRep.Open sqlRep, g_strCONN, 1, 3
		Do Until rsRep.EOF
			tmpName = rsRep("last name") & ", " & rsRep("first name")
			kulay = "#FFFFFF"
			If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
			strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'>" & tmpName & "</td></tr>"
			CSVBody = CSVBody & rsRep("last name") & "," & rsRep("first name") & vbCrLf
			If rsRep("ssnum") <> "" Then
				strBody = strBody & "<tr><td>&nbsp;</td><td class='tblgrn2'>Social Security</td>" & _
					"<td class='tblgrn2'>" & rsRep("ssnum") & "</td></tr>"
				CSVBody = CSVBody & "Social Security" & "," & rsRep("ssnum") & vbCrLf
			End If
			If rsRep("Passnum") <> "" Then
				strBody = strBody & "<tr><td>&nbsp;</td><td class='tblgrn2'>Passport</td>" & _
					"<td class='tblgrn2'>" & rsRep("Passnum") & "</td>" & _
					"<td class='tblgrn2'>" & rsRep("Passexp") & "</td></tr>"
				CSVBody = CSVBody & "Passport" & "," & rsRep("Passnum") & "," & rsRep("Passexp") & vbCrLf
			End If
			If rsRep("drivenum") <> "" Then
				strBody = strBody & "<tr><td>&nbsp;</td><td class='tblgrn2'>Driver's License</td>" & _
					"<td class='tblgrn2'>" & rsRep("drivenum") & "</td>" & _
					"<td class='tblgrn2'>" & rsRep("driveexp") & "</td></tr>"
				CSVBody = CSVBody & "Driver's License" & "," & rsRep("Drivenum") & "," & rsRep("Driveexp") & vbCrLf
			End If
			If rsRep("employnum") <> "" Then
				strBody = strBody & "<tr><td>&nbsp;</td><td class='tblgrn2'>Employment Authorization</td>" & _
					"<td class='tblgrn2'>" & rsRep("employnum") & "</td>" & _
					"<td class='tblgrn2'>" & rsRep("employexp") & "</td></tr>"
				CSVBody = CSVBody & "Employment Authorization" & "," & rsRep("employnum") & "," & rsRep("employexp") & vbCrLf
			End If
			If rsRep("greennum") <> "" Then
				strBody = strBody & "<tr><td>&nbsp;</td><td class='tblgrn2'>Green Card</td>" & _
					"<td class='tblgrn2'>" & rsRep("greennum") & "</td>" & _
					"<td class='tblgrn2'>" & rsRep("greenexp") & "</td></tr>"
				CSVBody = CSVBody & "Green Card" & "," & rsRep("greennum") & "," & rsRep("greenexp") & vbCrLf
			End If
			If rsRep("carnum") <> "" Then
				strBody = strBody & "<tr><td>&nbsp;</td><td class='tblgrn2'>Car Insurance</td>" & _
					"<td class='tblgrn2'>" & rsRep("carnum") & "</td>" & _
					"<td class='tblgrn2'>" & rsRep("carexp") & "</td></tr>"
				CSVBody = CSVBody & "Car Insurance" & "," & rsRep("carnum") & "," & rsRep("carexp") & vbCrLf
			End If
			'strBody = strBody & "</tr>" & vbCrlf
			rsRep.MoveNext
		Loop
		rsRep.Close
		Set rsRep = Nothing
	ElseIf Request("selRep") = 4 Then'hire 
		RepCSV =  "IntrHiredDate" & tmpdate & ".csv" 
		strMSG = "Interpreter Date of Hire" 
		strHead = "<td class='tblgrn'>Name</td>" & vbCrlf & _
			"<td class='tblgrn'>Date Of Hire</td>" & vbCrlf & vbCrlf 
		CSVHead = "Last Name, First Name, Date of Hire"
		Set rsRep = Server.CreateObject("ADODB.RecordSet")
		sqlRep = "SELECT * FROM Interpreter_T WHERE Active = 1" 
		If Request("txtRepFrom") <> "" And Request("txtRepTo") = "" Then
			sqlRep = sqlRep & " AND DateHired >= #" & Request("txtRepFrom") & "#"
		End If
		If Request("txtRepFrom") = "" And Request("txtRepTo") <> "" Then
			sqlRep = sqlRep & " AND DateHired <= #" & Request("txtRepTo") & "#"
		End If
		If Request("txtRepFrom") <> "" And Request("txtRepTo") <> "" Then
			sqlRep = sqlRep & " AND DateHired >= #" & Request("txtRepFrom") & "# AND DateHired <= #" & Request("txtRepTo") & "#"
		End If
		sqlRep = sqlRep & " ORDER BY [last name], [first name]"
		rsRep.Open sqlRep, g_strCONN, 1, 3
		Do Until rsRep.EOF
			tmpName = rsRep("last name") & ", " & rsRep("first name")
			kulay = "#FFFFFF"
			If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
			strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & tmpName & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("DateHired") & "</td></tr>" & vbCrLf
			CSVBody = CSVBody & rsRep("last name") & "," & rsRep("first name") & "," & rsRep("datehired") & vbCrLf
			y = y + 1
			rsRep.MoveNext
		Loop
		rsRep.Close
		Set rsRep = Nothing		
	ElseIf Request("selRep") = 5 Then'driver and crime
		RepCSV =  "IntrDriveCrime" & tmpdate & ".csv" 
		strMSG = "Interpreter Driver and Criminal Check" 
		strHead = "<td class='tblgrn'>Name</td>" & vbCrlf & _
			"<td class='tblgrn'>Driver Record</td>" & vbCrlf & _
			"<td class='tblgrn'>Criminal Record</td>" & vbCrlf 
		CSVHead = "Last Name, First Name, Document, Number, Expiration Date"
		Set rsRep = Server.CreateObject("ADODB.RecordSet")
		sqlRep = "SELECT * FROM Interpreter_T WHERE Active = 1 ORDER BY [last name], [first name]"
		rsRep.Open sqlRep, g_strCONN, 1, 3
		Do Until rsRep.EOF
			tmpName = rsRep("last name") & ", " & rsRep("first name")
			kulay = "#FFFFFF"
			If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
			strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & tmpName & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("DriveDate") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("CrimeDate") & "</td></tr>" & vbCrLf
			CSVBody = CSVBody & rsRep("last name") & "," & rsRep("first name") & "," & rsRep("drivedate") & "," & rsRep("crimedate") & vbCrLf
			y = y + 1
			rsRep.MoveNext
		Loop
		rsRep.Close
		Set rsRep = Nothing	
	ElseIf Request("selRep") = 6 Then'term 
		RepCSV =  "IntrTermDate" & tmpdate & ".csv" 
		strMSG = "Interpreter Date of Termination" 
		strHead = "<td class='tblgrn'>Name</td>" & vbCrlf & _
			"<td class='tblgrn'>Date Of Termination</td>" & vbCrlf & vbCrlf 
		CSVHead = "Last Name, First Name, Date of Termination"
		Set rsRep = Server.CreateObject("ADODB.RecordSet")
		sqlRep = "SELECT * FROM Interpreter_T WHERE Active = 0 " 
		If Request("txtRepFrom") <> "" And Request("txtRepTo") = "" Then
			sqlRep = sqlRep & " AND dateTerm >= #" & Request("txtRepFrom") & "#"
		End If
		If Request("txtRepFrom") = "" And Request("txtRepTo") <> "" Then
			sqlRep = sqlRep & " AND dateTerm <= #" & Request("txtRepTo") & "#"
		End If
		If Request("txtRepFrom") <> "" And Request("txtRepTo") <> "" Then
			sqlRep = sqlRep & " AND dateTerm >= #" & Request("txtRepFrom") & "# AND dateTerm <= #" & Request("txtRepTo") & "#"
		End If
		sqlRep = sqlRep & " ORDER BY [last name], [first name]"
		rsRep.Open sqlRep, g_strCONN, 1, 3
		Do Until rsRep.EOF
			tmpName = rsRep("last name") & ", " & rsRep("first name")
			kulay = "#FFFFFF"
			If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
			strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & tmpName & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("dateTerm") & "</td></tr>" & vbCrLf
			CSVBody = CSVBody & rsRep("last name") & "," & rsRep("first name") & "," & rsRep("dateterm") & vbCrLf
			y = y + 1
			rsRep.MoveNext
		Loop
		rsRep.Close
		Set rsRep = Nothing		
	End If
End If
If Request("csv") <> 1 Then
	'CONVERT TO CSV
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set Prt = fso.CreateTextFile(RepPath &  RepCSV, True)
	Prt.WriteLine "LANGUAGE BANK - REPORT"
	Prt.WriteLine strMSG
	Prt.WriteLine CSVHead
	Prt.WriteLine CSVBody
	Prt.Close	
	Set Prt = Nothing
	
	'COPY FILE TO BACKUP
	
	fso.CopyFile RepPath & RepCSV, BackupStr
	'If Request("bill") = 1 Then
	'	Set Prt2 = fso.CreateTextFile(RepPath2 & RepCSV2, True)
	'	Prt2.WriteLine "LANGUAGE BANK - REPORT"
	'	Prt2.WriteLine strMSG2
	'	Prt2.WriteLine CSVHead2
	'	Prt2.WriteLine CSVBody2
	'	Prt2.Close	
	'	Set Prt2 = Nothing
		'COPY FILE TO BACKUP
	'	fso.CopyFile RepPath2 &  RepCSV2, BackupStr
	'End If
	Set fso = Nothing
	'EXPORT CSV
	'If Request("bill") <> 1 Then
		tmpstring = "CSV/" & repCSV
	'Else
	'	tmpstring= "CSV/" & repCSV2
	'End IF
Else
	'EXPORT CSV
	
	'Set dload = Server.CreateObject("SCUpload.Upload")

	'If Request("bill") <> 1 Then
		'dload.Download RepCSV
	'	tmpstring = "CSV/InstBillReq262007.csv"
	'Else
		'tmpstring= "CSV/IntrBillReq262007.csv"
	'	'dload.Download RepCSV2
	'End IF
	'Set dload = Nothing
End If
%>
<html>
	<head>
		<title>Language Bank - Expiring Documents Report</title>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function exportMe()
		{
			document.frmResult.action = "printreport.asp?csv=1"
			document.frmResult.submit();
		}
		function PassMe(xxx)
		{
			window.opener.document.frmReport.hideID.value = xxx;
			window.opener.SubmitAko();
			self.close();
		}
		-->
		</script>
	</head>
	<body>
		<form method='post' name='frmResult'>
			<table cellSpacing='0' cellPadding='0' width="100%" bgColor='white' border='0'>
				<tr>
					<td valign='top'>
						<table bgColor='white' border='0' cellSpacing='0' cellPadding='0' align='center'>
						<tr>
							<td>
								<img src='images/LBISLOGO.jpg' align='center'>
							</td>
						</tr>
						<tr>
							<td align='center'>
								261&nbsp;Sheep&nbsp;Davis&nbsp;Road,&nbsp;Concord,&nbsp;NH&nbsp;03301<br>
								Tel:&nbsp;(603)&nbsp;410-6183&nbsp;|&nbsp;Fax:&nbsp;(603)&nbsp;410-6186
							</td>
						</tr>
						</table>
					</td>
				</tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td valign='top' >
						<table bgColor='white' border='0' cellSpacing='2' cellPadding='0' align='center'>
							<tr bgcolor='#C2AB4B'>
								<td colspan='4' align='center'>
									
										<b><%=strMSG%></b>
									
								</td>
							</tr>
							<tr>
								
								<%=strHead%>
							
							</tr>
							
								<%=strBody%>
							
							<tr><td>&nbsp;</td></tr>
							<tr>
								<td colspan='4' align='center' height='100px' valign='bottom'>
									<input class='btn' type='button' value='Print' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='print({bShrinkToFit: true});'>
									<%'<input class='btn' type='button' value='CSV Export' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='exportMe();'>%>
									<input class='btn' type='button' value='CSV Export' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="document.location='<%=tmpstring%>';">
								</td>
							</tr>
								<td colspan='4' align='center' height='100px' valign='bottom'>
									* If needed, please adjust the page orientation of your printer to landscape to view all columns in a single page   
								</td>
							<tr>
							</tr>
						</table>	
					</td>
				</tr>
			</table>
		</form>
	</body>
</html>