<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%
If Request.Cookies("LBUSERTYPE") <> 1 Then 
	Session("MSG") = "Invalid account."
	Response.Redirect "default.asp"
End If
Function Z_FormatTime(xxx)
	Z_FormatTime = Null
	If xxx <> "" Or Not IsNull(xxx)  Then
		If IsDate(xxx) Then Z_FormatTime = FormatDateTime(xxx, 4) 
	End If
End Function
'USER CHECK
If Cint(Request.Cookies("LBUSERTYPE")) = 2 Then
	Session("MSG") = "Error: Invalid user type. Please sign-in again."
	Response.Redirect "default.asp"
End If
server.scripttimeout = 360000
Function MyStatus(xxx)
	Select Case xxx
		Case 1
			MyStatus = "<font color='#000000' size='+3'>•</font>"
		Case 2
			MyStatus = "<font color='#0000FF' size='+3'>•</font>"
		Case 3
			MyStatus = "<font color='#FF0000' size='+3'>•</font>"
		Case 4
			MyStatus = "<font color='#FF00FF' size='+3'>•</font>"
		Case Else
			MyStatus = ""
	End Select
End Function
Function GetMyDept(xxx)
	GetMyDept = ""
	Set rsDept = Server.CreateObject("ADODB.RecordSet")
	sqlDept = " SELECT Dept FROM dept_T WHERE [index] = " & xxx
	rsDept.Open sqlDept, g_strCONN, 3, 1
	If Not rsDept.EOF Then
		GetMyDept = " - " & rsDept("Dept")
	End If
	rsDept.Close
	Set rsDept = Nothing
End Function
tmpPage = "document.frmTbl."
radioApp = ""
radioID = ""
radioAll = "checked"
radioAss = "checked"
radioUnass = ""
radioUnass2 = ""
If Request("ctrlX") = 1 Then
	mybtn = "Save Timesheet"
Else
	mybtn = "Save Mileage"
End If
x = 0
If Request.ServerVariables("REQUEST_METHOD") = "POST"  Or Request("action") = 3 Then

		sqlReq = "SELECT overmile, payhrs, overpayhrs, request_T.InstID, IntrID, LangID, InstRate, SentReq, Processed, Status, DeptID, request_T.[index]," & _
			" clname, cfname, appDate, actTT, actMil, astarttime, aendtime, toll, LBconfirm, LbconfirmToll " & _
			"FROM request_T, institution_T, language_T, interpreter_T, requester_T, dept_T " & _
			"WHERE request_T.InstID = institution_T.[index] " & _
			"AND LangId = language_T.[index] " & _
			"AND IntrId = interpreter_T.[index] " & _
			"AND request_T.DeptId = dept_T.[index] " & _
			"AND ReqID = requester_T.[index] " & _
			"AND showintr = 1 " & _
			"AND status <> 2 AND status <> 3 " 
		If Request("ctrlX") = 1 Then
			If Request("radioAss") = 0 Then	
				sqlReq = sqlReq & "AND lbconfirm = 0 "
				radioAss = "checked"
				radioUnass = ""
				radioUnass2 = ""
			ElseIf Request("radioAss") = 1 Then	
				sqlReq = sqlReq & "AND lbconfirm = 1 "
				radioAss = ""
				radioUnass = "checked"
				radioUnass2 = ""
			Else
				radioAss = ""
				radioUnass = ""
				radioUnass2 = "checked"
			End If
		Else
			If Request("radioAss") = 0 Then	
				sqlReq = sqlReq & "AND LbconfirmToll = 0 "
				radioAss = "checked"
				radioUnass = ""
				radioUnass2 = ""
			ElseIf Request("radioAss") = 1 Then	
				sqlReq = sqlReq & "AND LbconfirmToll = 1 "
				radioAss = ""
				radioUnass = "checked"
				radioUnass2 = ""
			Else
				radioAss = ""
				radioUnass = ""
				radioUnass2 = "checked"
			End If
		End If
	
	'FIND
	If Request("radioStat") = 0 Then
		radioApp = "checked"
		radioID = ""
		radioAll = ""
		If Request("txtFromd8") <> "" Then
			If IsDate(Request("txtFromd8")) Then
				sqlReq = sqlReq & " AND appDate >= '" & Request("txtFromd8") & "' "
				tmpFromd8 = Request("txtFromd8") 
			Else
				Session("MSG") = "ERROR: Invalid Appointment Date Range (From)."
				Response.Redirect "reqtable2.asp"
			End If
		End If
		If Request("txtTod8") <> "" Then
			If IsDate(Request("txtTod8")) Then
				sqlReq = sqlReq & " AND appDate <= '" & Request("txtTod8") & "' "
				tmpTod8 = Request("txtTod8")
			Else
				Session("MSG") = "ERROR: Invalid Appointment Date Range (To)."
				Response.Redirect "reqtable2.asp"
			End If
		End If
	ElseIf Request("radioStat") = 1 Then
		radioApp = ""
		radioID = "checked"
		radioAll = ""
		If Request("txtFromID") <> "" Then
			If IsNumeric(Request("txtFromID")) Then
				sqlReq = sqlReq & " AND request_T.[index] >= " & Request("txtFromID")
				tmpFromID = Request("txtFromID")
			Else
				Session("MSG") = "ERROR: Invalid Appointment ID Range (From)."
				Response.Redirect "reqtable2.asp"
			End If
		End If
		If Request("txtToID") <> "" Then
			If IsNumeric(Request("txtToID")) Then
				sqlReq = sqlReq & " AND request_T.[index] <= " & Request("txtToID")
				tmpToID = Request("txtToID")
			Else
				Session("MSG") = "ERROR: Invalid Appointment ID Range (To)."
				Response.Redirect "reqtable2.asp"
			End If
		End If
	Else
		radioApp = ""
		radioID = ""
		radioAll = "checked"
	End If
	'FILTER
	xInst = Cint(Request("selInst"))
	If xInst <> -1 Then 
		sqlReq = sqlReq & " AND "
		sqlReq = sqlReq & "request_T.InstID = " & xInst
	End If
	xLang = Cint(Request("selLang"))
	If xLang <> -1 Then 
		sqlReq = sqlReq & " AND "
		sqlReq = sqlReq & "LangID = " & xLang
	End If
	If Cint(Request.Cookies("LBUSERTYPE")) <> 4 Then
			If Trim(Request("txtclilname")) <> "" Then
				sqlReq = sqlReq & " AND Upper(Clname) LIKE '" & Ucase(Trim(Request("txtclilname"))) & "%'"
			End If
			If Trim(Request("txtclifname")) <> "" Then
				sqlReq = sqlReq & " AND Upper(Cfname) LIKE '" & Ucase(Trim(Request("txtclifname"))) & "%'"
			End If

	End If
	xIntr = Cint(Request("selIntr"))
	If xIntr <> -1 Then 
		sqlReq = sqlReq & " AND "
		sqlReq = sqlReq & "IntrID = " & xIntr
	End If
	xClass = Cint(Request("selClass"))
	If xClass <> -1 Then 
		sqlReq = sqlReq & " AND "
		sqlReq = sqlReq & "Class = " & xClass
	End If
	'ADMIN ONLY
	xAdmin = Z_CZero(Request("selAdmin"))
	If xAdmin = 1 Then
		sqlReq = sqlReq & " AND (Status = 1) AND Processed IS NULL"
		meUnBilled = "selected"
	ElseIf xAdmin = 2 Then
		sqlReq = sqlReq & " AND (Status = 1 OR Status = 4) AND NOT Processed IS NULL"
		meBilled = "selected"
	ElseIf xAdmin = 3 Then
		sqlReq = sqlReq & " AND (Status = 2)"
		meMisded = "selected"
	ElseIf xAdmin = 4 Then
		sqlReq = sqlReq & " AND (Status = 3)"
		meCanceled = "selected"
	ElseIf xAdmin = 5 Then
		sqlReq = sqlReq & " AND (Status = 4)"
		meCanceledBill = "selected"
	ElseIf xAdmin = 6 Then
		sqlReq = sqlReq & " AND (Status = 0)"
		mePending = "selected"
	Else
		'sqlReq = sqlReq & " AND IsNull(Processed)"
	End If
	sqlReq = sqlReq & " ORDER BY appDate, Facility, [last name], [first name]"
'End If
'GET REQUESTS
'response.write sqlReq
Set rsReq = Server.CreateObject("ADODB.RecordSet")
rsReq.Open sqlReq, g_strCONN, 3, 1
x = 1
If Not rsReq.EOF Then
	Do Until rsReq.EOF
		kulay = ""
		If Not Z_IsOdd(x) Then kulay = "#FBEEB7"
		'GET INSTITUTION
		Set rsInst = Server.CreateObject("ADODB.RecordSet")
		sqlInst = "SELECT Facility FROM institution_T WHERE [index] = " & rsReq("InstID")
		rsInst.Open sqlInst, g_strCONN, 3, 1
		If Not rsInst.EOF Then
			tmpIname = rsInst("Facility")  
			'If rsInst("Department") <> "" Then tmpIname = tmpIname & " <br> " & rsInst("Department")
		Else
			tmpIname = "N/A"
		End If
		rsInst.Close
		Set rsInst = Nothing 
		'GET INTERPRETER INFO
		Set rsIntr = Server.CreateObject("ADODB.RecordSet")
		sqlIntr = "SELECT [last name], [first name] FROM interpreter_T WHERE [index] = " & rsReq("IntrID")
		rsIntr.Open sqlIntr, g_strCONN, 3, 1
		If Not rsIntr.EOF Then
			tmpInName = rsIntr("last name") & ", " & rsIntr("first name")
		Else
			tmpInName = "N/A"
		End If
		rsIntr.Close
		Set rsIntr = Nothing
		'GET LANGUAGE
		Set rsLang = Server.CreateObject("ADODB.RecordSet")
		sqlLang  = "SELECT [language] FROM language_T WHERE [index] = " & rsReq("LangID")
		rsLang.Open sqlLang , g_strCONN, 3, 1
		If Not rsLang.EOF Then
			tmpSalita = rsLang("language") 
		Else
			tmpSalita = "N/A"
		End If
		rsLang.Close
		Set rsLang = Nothing 
	
		Stat = MyStatus(rsReq("Status") )
		myDept =  GetMyDept(rsReq("DeptID"))
		TT = Z_FormatNumber(rsReq("actTT"), 2)
		If rsReq("overpayhrs") Then 
			BlnOver = "checked"
			PHrs = Z_FormatNumber(rsReq("payhrs"), 2)
		Else
			BlnOver = ""
			PHrs = Z_FormatNumber(IntrBillHrs(rsReq("AStarttime"), rsReq("AEndtime")), 2)
		End If
		FPHrs = Z_Czero(PHrs) + Z_Czero(TT)
		BlnOver2 = ""
		If rsReq("overmile") Then BlnOver2 = "checked"
		tmpAMT = Z_FormatNumber(rsReq("actMil"), 2)
		LBcon = ""
		LBconx = ""
		LBconxx = ""
		If rsReq("LBconfirm") = True Then 
			LBcon = "Checked disabled"
			LBconx = "readonly"
			LBconxx = "disabled"
		End If
		LBcon2 = ""
		LBconx2 = ""
		LBconxx2 = ""
		If rsReq("LBconfirmToll") = True Then 
			LBcon2 = "Checked disabled"
			LBconx2 = "readonly"
			LBconxx2 = "disabled"
		End If
			strtbl = strtbl & "<tr bgcolor='" & kulay & "'>" & vbCrLf & _ 
				"<td class='tblgrn2' width='10px'>" & Stat & "</td>" & vbCrLf & _
				"<td class='tblgrn2' ><input type='hidden' name='ID" & x & "' value='" & rsReq("Index") & "'><a class='link2' href='reqconfirm.asp?ID=" & rsReq("Index") & "'><b>" & rsReq("Index") & "</b></a></td>" & vbCrLf & _
				"<td class='tblgrn2' ><nobr>" & tmpIname & myDept & "</td>" & vbCrLf & _
				"<td class='tblgrn2' >" & tmpSalita & "</td>" & vbCrLf & _
				"<td class='tblgrn2' >" & rsReq("clname") & ", " & rsReq("cfname") & "</td>" & vbCrLf & _
				"<td class='tblgrn2' >" & tmpInName & "</td>" & vbCrLf & _
				"<td class='tblgrn2' >" & rsReq("appDate") & "</td>" & vbCrLf
				If Request("ctrlX") = 1 Then
					strtbl = strtbl & "<td class='tblgrn2' >" & TT & "</td>" & vbCrLf & _
					"<td class='tblgrn2' ><input class='main2' name='txtstime" & x & "' maxlength='5' size='7' " & LBconx & " value='" & Z_FormatTime(rsReq("astarttime")) & "' onKeyUp=""javascript:return maskMe(this.value,this,'2,6',':');"" onBlur=""javascript:return maskMe(this.value,this,'2,6',':');""></td>" & vbCrLf & _
					"<td class='tblgrn2' ><input class='main2' name='txtetime" & x & "' maxlength='5' size='7' " & LBconx & " value='" & Z_FormatTime(rsReq("aendtime")) & "' onKeyUp=""javascript:return maskMe(this.value,this,'2,6',':');"" onBlur=""javascript:return maskMe(this.value,this,'2,6',':');""></td>" & vbCrLf & _
					"<td class='tblgrn2' ><nobr><input class='main2' name='txtPhrs" & x & "' maxlength='6' size='7' " & LBconx & " value='" & PHrs & "'><input type='checkbox' name='chkOverPhrs" & x & "' value='1' " & LBconxx & " " & BlnOver & " ></td>" & vbCrLf & _
					"<td class='tblgrn2' >" & FPHrs & "</td>" & vbCrLf & _
					"<td class='tblgrn2' ><input type='checkbox' ID='chkTS" & x & "' name='chkTS" & x & "' value='1' " & LBcon & "></td>" & vbCrLf
				Else
					strtbl = strtbl & "<td class='tblgrn2' ><nobr><input class='main2' name='txtmile" & x & "' maxlength='6' size='7' " & LBconx2 & " value='" & tmpAMT & "'><input type='checkbox' name='chkOverMile" & x & "' value='1' " & LBconxx2 & " " & BlnOver2 & " ></td>" & vbCrLf & _
					"<td class='tblgrn2' ><nobr>$<input class='main2' name='txtTol" & x & "' maxlength='5' size='7' " & LBconx2 & " value='" & Z_FormatNumber(rsReq("toll"), 2) & "'></td>" & vbCrLf & _
					"<td class='tblgrn2' ><input type='checkbox' ID='chkM" & x & "' name='chkM" & x & "' value='1' " & LBcon2 & "></td></tr>" & vbCrLf
				End If
			strtbl = strtbl & "</tr>" & vbCrLf
		x = x + 1
		rsReq.MoveNext
	Loop
Else
	strtbl = "<tr><td colspan='14' align='center'><i>&lt -- No records found. -- &gt</i></td></tr>"
End If
rsReq.Close
Set rsReq = Nothing
End If
'SORT
If Request("sType") <> "" Then
	If Request("stype") = 1 Then stype = 2
	If Request("stype") = 2 Then stype = 1
Else
	stype = 1
End If
'FILTER CRITERIA
tmpclilname = Request("txtclilname")
tmpclifname = Request("txtclifname")
'GET INSTITUTION LIST
Set rsInst = Server.CreateObject("ADODB.RecordSet")
sqlInst = "SELECT Facility, [Index] FROM institution_T ORDER BY [Facility]"
rsInst.Open sqlInst, g_strCONN, 3, 1
Do Until rsInst.EOF
	InstSel = ""
	If Cint(Request("selInst")) = rsInst("Index") Then InstSel = "selected"
	InstName = rsInst("Facility")
	strInst = strInst	& "<option value='" & rsInst("Index") & "' " & InstSel & ">" &  InstName & "</option>" & vbCrlf
	rsInst.MoveNext
Loop
rsInst.Close
Set rsInst = Nothing
'GET AVAILABLE LANGUAGES
Set rsLang = Server.CreateObject("ADODB.RecordSet")
sqlLang = "SELECT [Index], [language] FROM language_T ORDER BY [Language]"
rsLang.Open sqlLang, g_strCONN, 3, 1
Do Until rsLang.EOF
	LangSel = ""
	If Cint(Request("selLang")) = rsLang("Index") Then LangSel = "selected"
	strLang = strLang	& "<option value='" & rsLang("Index") & "' " & LangSel & ">" &  rsLang("language") & "</option>" & vbCrlf
	rsLang.MoveNext
Loop
rsLang.Close
Set rsLang = Nothing
'GET INTERPRETER LIST
Set rsIntr = Server.CreateObject("ADODB.RecordSet")
sqlIntr = "SELECT [Index], [last name], [first name] FROM interpreter_T WHERE Active = 1 ORDER BY [last name], [first name]"
rsIntr.Open sqlIntr, g_strCONN, 3, 1
Do Until rsIntr.EOF
	IntrSel = ""
	If Cint(Request("selIntr")) = rsIntr("Index") Then IntrSel = "selected"
	strIntr = strIntr	& "<option value='" & rsIntr("Index") & "' " & IntrSel & ">" & rsIntr("last name") & ", " & rsIntr("first name") & "</option>" & vbCrlf
	rsIntr.MoveNext
Loop
rsIntr.Close
Set rsIntr = Nothing
If Cint(Request.Cookies("LBUSERTYPE")) <> 4 Then 
	
End If
'FOR CLASSIFICATION
tmpClass = Cint(Request("selClass"))
Select Case tmpClass
	Case 1 SocSer = "selected"
    Case 2 Priv = "selected"
	Case 3 Legal = "selected"	
	Case 4 Med = "selected"
End Select
'FOR ADMIN
tmpAdmin = Z_CZero(Request("selAdmin"))
Select Case tmpAdmin
	Case 1 meUnBilled = "selected"
    Case 2 meBilled = "selected"
	Case 3 meMisded = "selected"
	Case 4 meCanceled = "selected"
	Case 5 meCanceledBill = "selected"
End Select
%>
<html>
	<head>
		<title>Language Bank - Timesheet/Mileage</title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function SaveMe()
		{
			var ans = window.confirm("This action will save all entries inside the table to the database. Please double check your enties.\nClick Cancel to stop.");
			if (ans)
			{
				document.frmTbl.action = "action.asp?ctrl=17";
				document.frmTbl.submit();
			}
		}
		function SortMe(sortnum)
		{
			document.frmTbl.action = "reqtable2.asp?sort=" + sortnum + "&sType=" + <%=stype%>;
			document.frmTbl.submit();
		}
		function FindMe(xxx)
		{
			document.frmTbl.action = "reqtable2.asp?ctrlX=" + xxx;
			document.frmTbl.submit();
		}
		function FixSort()
		{
			document.frmTbl.txtFromd8.disabled = true;
			document.frmTbl.txtTod8.disabled = true;
			document.frmTbl.txtFromID.disabled = true;
			document.frmTbl.txtToID.disabled = true;
			if (document.frmTbl.radioStat[0].checked == true)
			{
				document.frmTbl.txtFromd8.disabled = false;
				document.frmTbl.txtTod8.disabled = false;
			}
			if (document.frmTbl.radioStat[1].checked == true)
			{
				document.frmTbl.txtFromID.disabled = false;
				document.frmTbl.txtToID.disabled = false;
			}
		}
		function CalendarView(strDate)
		{
			document.frmTbl.action = 'calendarview2.asp?appDate=' + strDate;
			document.frmTbl.submit();
		}
		function maskMe(str,textbox,loc,delim)
		{
			var locs = loc.split(',');
			for (var i = 0; i <= locs.length; i++)
			{
				for (var k = 0; k <= str.length; k++)
				{
					 if (k == locs[i])
					 {
						if (str.substring(k, k+1) != delim)
					 	{
					 		str = str.substring(0,k) + delim + str.substring(k,str.length);
		     			}
					}
				}
		 	}
			textbox.value = str
		}
		function checkme(xxx)
		{
			var tmpElem;
			var z;
			if (document.frmTbl.chkall.checked == true)
			{
				for(z = 1; z <= xxx; z ++)
				{
					<% If Request("ctrlX") = 1 Then %>
						tmpElem = "chkTS" + z;
					<% Else %>
						tmpElem = "chkM" + z;
					<% End If %>
					document.getElementById(tmpElem).checked = true;
				}	
			}
			else
			{
				for(z = 1; z <= xxx; z ++)
				{
					<% If Request("ctrlX") = 1 Then %>
						tmpElem = "chkTS" + z;
					<% Else %>
						tmpElem = "chkM" + z;
					<% End If %>
					document.getElementById(tmpElem).checked = false;
				}	
			}
		}
		-->
		</script>
		<style type="text/css">
	 	.container
	      {
	          border: solid 1px black;
	          overflow: auto;
	      }
	      .noscroll
	      {
	          position: relative;
	          background-color: white;
	          top:expression(this.offsetParent.scrollTop);
	      }
	      th
	      {
	          text-align: left;
	      }
		</style>
		<body onload='FixSort();'>
			<form method='post' name='frmTbl' action='reqtable2.asp'>
				<table cellSpacing='0' cellPadding='0' height='100%' width="100%" border='0' class='bgstyle2'>
					<tr>
						<td height='100px'>
							<!-- #include file="_header.asp" -->
						</td>
					</tr>
					<tr>
						<td valign='top'>
							<table cellSpacing='0' cellPadding='0' width="100%" border='0'>
									<!-- #include file="_greetme.asp" -->
								<tr>
									<td>
										<table cellpadding='0' cellspacing='0' width='100%' border='0'>
											<tr>
												<td align='left'>
													Legend: <font color='#000000' size='+3'>•</font>&nbsp;-&nbsp;completed&nbsp;<font color='#0000FF' size='+3'>•</font>&nbsp;-&nbsp;missed&nbsp;<font color='#FF0000 ' size='+3'>•</font>&nbsp;-&nbsp;Canceled&nbsp;
													<font color='#FF00FF' size='+3'>•</font>&nbsp;-&nbsp;Canceled (billable)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
													<% If Cint(Request.Cookies("LBUSERTYPE")) = 1 Or Cint(Request.Cookies("LBUSERTYPE")) = 3 Or Cint(Request.Cookies("LBUSERTYPE")) = 0 Then %>
														Admin Sort:
														<select class='seltxt' style='width:100px;' name='selAdmin'>
															<option value='0'>&nbsp;</option>
															<option <%=mePending%> value='6'>Pending</option>
															<option <%=meUnBIlled%> value='1'>Completed (Unbilled)</option>
															<option <%=meCanceledBill%> value='5'>Canceled (Billable)</option>
															<option <%=meBilled%> value='2'>BILLED</option>
														</select>
														<input class='btntbl' type='button' value='GO' onmouseover="this.className='hovbtntbl'" onmouseout="this.className='btntbl'" onclick='FindMe(<%=Request("ctrlX")%>);'>
													<% End If %>
												</td>
												<% If Cint(Request.Cookies("LBUSERTYPE")) <> 4 Then %>
													<% If Cint(Request.Cookies("LBUSERTYPE")) <> 1 Then %> 
														<td align='right'>
															<input type='hidden' name='Hctr' value='<%=x%>'>
															<input class='btntbl' type='button' value='Save Table' style='height: 25px; width: 200px;' onmouseover="this.className='hovbtntbl'" onmouseout="this.className='btntbl'" onclick='SaveMe();'>
														</td>
													<% Else %>
													<td align='right'>
															<input type='hidden' name='ctrlX' value='<%=Request("ctrlX")%>'>
															<input type='hidden' name='Hctr' value='<%=x%>'>
															<input class='btntbl' type='button' value='<%=mybtn%>' style='height: 25px; width: 150px;' onmouseover="this.className='hovbtntbl'" onmouseout="this.className='btntbl'" onclick='SaveMe();'>
														</td>
													<% End If %>
												<% Else %>
													<td>&nbsp;</td>
												<% End If %>
											</tr>
										</table>
									</td>
								</tr>
								<% If Session("MSG") <> "" Then %>
									<tr><td>&nbsp;</td></tr>
									<tr>
										<td colspan='14' align='left'>
											<div name="dErr" style="width:300px; height:40px;OVERFLOW: auto;">
												<table border='0' cellspacing='1'>		
													<tr>
														<td><span class='error'><%=Session("MSG")%></span></td>
													</tr>
												</table>
											</div>
										</td>
									</tr>
									<tr><td>&nbsp;</td></tr>
								<% End If %>
								<tr>
									<td colspan='10' align='left'>
										<div class='container' style='height: 500px; width:1000px; position: relative;'>
											<table class="reqtble" width='100%'>	
												<thead>
													<tr class="noscroll">	
														<td colspan='2' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'" class='tblgrn'>Request ID</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Institution</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Language</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Client</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Interpreter</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Appointment Date</td>
														<% If Request("ctrlX") = 1 Then %>
															<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Travel Time</td>
															<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Actual Start Time</td>
															<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Actual End Time</td>
															<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Payable Hours</td>
															<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Final Payable Hours</td>
															<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">
																Approve Timesheet<br>
																<input type='checkbox' name='chkall' onclick='checkme(<%=x%>);'>
															</td>
														<% Else %>
															<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Mileage</td>
															<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Tolls & parking</td>
															<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">
																Approve Mileage<br>
																<input type='checkbox' name='chkall' onclick='checkme(<%=x%>);'>
															</td>
														<% End If %>
													</tr>
												</thead>
												<tbody style="OVERFLOW: auto;">
													<%=strtbl%>
												</tbody>
											</table>
										</div>	
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td>
							<table width='100%'  border='0'>
								<tr>
									<td align='left'>
										&nbsp;
									</td>
									<td align='right'>
										<% If x <> 0 Then %>
											<b><u><%=x - 1%></u></b> records &nbsp;&nbsp;&nbsp;&nbsp;
										<% End If %>
									</td>
									<td>&nbsp;</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td>
							<table cellSpacing='0' cellPadding='0' width='1005px' border='0' style='border: solid 1px;'>
								<tr bgcolor='#FBEEB7'>
									<td align='right' style='border-bottom: solid 1px;'><b>Sort:</b></td>
									<td style='border-right: solid 1px;border-bottom: solid 1px;'>
										<input type='radio' name='radioStat' value='0' <%=radioApp%> onclick='FixSort();'>&nbsp;<b>App. Date Range:</b>
										&nbsp;&nbsp;
										<input class='main' size='10' maxlength='10' name='txtFromd8' value='<%=tmpFromd8%>'>
										&nbsp;-&nbsp;
										<input class='main' size='10' maxlength='10' name='txtTod8' value='<%=tmpTod8%>'>
										<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">mm/dd/yyyy</span>
										&nbsp;&nbsp;
										<input type='radio' name='radioStat' value='1' <%=radioID%> onclick='FixSort();'>&nbsp;<b>Request ID Range:</b>
										&nbsp;&nbsp;
										<input class='main' size='7' maxlength='7' name='txtFromID' value='<%=tmpFromID%>'>
										&nbsp;-&nbsp;
										<input class='main' size='7' maxlength='7' name='txtToID' value='<%=tmpToID%>'>
										&nbsp;&nbsp;
										<input type='radio' name='radioStat' value='2' <%=radioAll%> onclick='FixSort();'>&nbsp;<b>All</b>
									</td>
									<td align='right' style='border-bottom: solid 1px;'><b>&nbsp;&nbsp;</b></td>
									<td style='border-bottom: solid 1px;'>
										<input type='radio' name='radioAss' value='0' <%=radioAss%> onclick='FixSort();'>&nbsp;<b>Unapproved</b>
										&nbsp;&nbsp;
										<input type='radio' name='radioAss' value='1' <%=radioUnAss%> onclick='FixSort();'>&nbsp;<b>Approved</b>
										&nbsp;&nbsp;
										<input type='radio' name='radioAss' value='2' <%=radioUnAss2%> onclick='FixSort();'>&nbsp;<b>ALL</b>
										&nbsp;&nbsp;
									</td>
									<td align='right' style='border-left: solid 1px;' rowspan='3'>
										<input class='btntbl' type='button' value='Find' style='height: 35px;' onmouseover="this.className='hovbtntbl'" onmouseout="this.className='btntbl'" onclick='FindMe(<%=Request("ctrlX")%>);'>
									</td>
									</td>
								</tr>
								<tr bgcolor='#FBEEB7'>
									<td align='left' colspan='4'>
										Institution:
										<select class='seltxt' style='width: 285px;' name='selInst'>
											<option value='-1'>&nbsp;</option>
											<%=strInst%>
										</select>
										&nbsp;Language:
										<select class='seltxt' style='width: 150px;' name='selLang'>
											<option value='-1'>&nbsp;</option>
											<%=strLang%>
										</select>
										<% If Cint(Request.Cookies("LBUSERTYPE")) <> 4 Then %>
											&nbsp;Client:
											<input class='main' size='20' maxlength='20' name='txtclilname' value='<%=tmpclilname%>'>
											&nbsp;,&nbsp;&nbsp;
											<input class='main' size='20' maxlength='20' name='txtclifname' value='<%=tmpclifname%>'>
											<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">Last name, First name</span>
										<% End If %>
										
										&nbsp;
									</td>
								</tr>
								<tr bgcolor='#FBEEB7'>
									<td align='left' colspan='4'>
										Interpreter:
										<select class='seltxt' name='selIntr'>
											<option value='-1'>&nbsp;</option>
											<%=strIntr%>
										</select>
										&nbsp;Classification:
										<select class='seltxt' style='width: 100px;' name='selClass'>
											<option value='-1'>&nbsp;</option>
											<option value='1' <%=SocSer%>>Social Services</option>
											<option value='2' <%=Priv%>>Private</option>
											<option value='3' <%=Legal%>>Legal</option>
											<option value='4' <%=Med%>>Medical</option>
										</select>
									</td>
									<td>&nbsp;</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td height='50px' valign='bottom'>
							<!-- #include file="_footer.asp" -->
						</td>
					</tr>
				</table>
			</form>
		</body>
	</head>
</html>
<%
If Session("MSG") <> "" Then
	tmpMSG = Replace(Session("MSG"), "<br>", "\n")
%>
<script><!--
	alert("<%=tmpMSG%>");
--></script>
<%
End If
Session("MSG") = ""
%>