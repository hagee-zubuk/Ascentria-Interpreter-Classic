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
	Function Z_FormatTime(xxx)
		Z_FormatTime = Null
		If xxx <> "" Or Not IsNull(xxx)  Then
			If IsDate(xxx) Then Z_FormatTime = FormatDateTime(xxx, 4) 
		End If
	End Function
	tmpPage = "document.frmTS."
	'Get week range
	If Request("action") = 1 Then
		sundate = DateAdd("d", -7, GetSun(Request("tmpDate")))
		satdate = DateAdd("d", -7, GetSat(Request("tmpDate")))
	ElseIf Request("action") = 2 Then
		sundate = DateAdd("d", 7, GetSun(Request("tmpDate")))
		satdate = DateAdd("d", 7, GetSat(Request("tmpDate")))
	Else 
		sundate = GetSun(Request("tmpDate"))
		satdate = GetSat(Request("tmpDate"))
	End If
	Set rsTS = Server.CreateObject("ADODB.RecordSet")
	sqlTS = "SELECT InstID, noreas, DSnoreas, status, happen, overpayhrs, payhrs, LBconfirm, [index], totalhrs, Status, confirmed, InstID, Cfname, Clname, AStarttime, AEndtime, totalhrs, actTT, appDate, toll, " & _
		"LbconfirmToll, deptID FROM request_T WHERE appDate >= '" & sundate & "' AND appDate <= '" & satDate & "' AND IntrID = " & Session("UIntr") & " " & _
		"AND showintr = 1 AND status <> 2 AND Status <> 3 ORDER BY appDate ,appTimeFrom"
	rsTS.Open sqlTS, g_strCONN, 1, 3
	ctr = 0
	ctrCon = 0
	If Not rsTS.EOF Then
		Do Until rsTS.EOF
			tmpTot = rsTS("totalhrs")
			myStat = ""
			'If rsTS("Status") = 2 Or rsTS("Status") = 3 Or rsTS("Status") = 4 Or rsTS("confirmed") <> "" Then myStat = "DISABLED"
			'tmpAMT = "$" & Z_FormatNumber(AmtRate(rsTS("m_intr")), 2)
			TT = Z_FormatNumber(rsTS("actTT"), 2)
			If rsTS("overpayhrs") Then 
				PHrs = Z_FormatNumber(rsTS("payhrs"), 2)
			Else
				PHrs = Z_FormatNumber(IntrBillHrs(rsTS("AStarttime"), rsTS("AEndtime")), 2)
			End If
			FPHrs = Z_Czero(PHrs) + Z_Czero(TT)
			TotFPHrs = TotFPHrs + FPHrs
			tmpCon = ""
			LBcon = ""
			If rsTS("LBconfirm") = True Then 
				tmpCon = "checked" '"<b>*</b>"
				LBcon = "readonly"
				ctrCon = ctrCon + 1
			End If
			tmpConToll = ""
			LBconToll = ""
			If rsTS("LbconfirmToll") = True Then 
				tmpConToll = "checked"
				LBconToll = "readonly"
			End If
			IntrCon = ""
			If rsTS("confirmed") <> "" Then IntrCon = "disabled checked"
			myAct = rsTS("index") & " - " & GetInst(rsTS("InstID")) & " - " & left(rsTS("Cfname"), 1) & ". " & left(rsTS("Clname"), 1) & "."
			AStime = Z_FormatTime(Ctime(rsTS("AStarttime")))
			AEtime = Z_FormatTime(Ctime(rsTS("AEndtime")))
			q0 = ""
			q1 = ""
			q2 = ""
			If rsTS("happen") = 0 Then q0 = "SELECTED"
			If rsTS("happen") = 1 Then q1 = "SELECTED"
			If rsTS("happen") = 2 Then q2 = "SELECTED"
			no0 = ""
			no1 = ""
			no2 = ""
			no3 = ""
			no7 = ""
			no8 = ""
			If rsTS("noreas") = 0 Then no0 = "SELECTED"
			If rsTS("noreas") = 1 Then no1 = "SELECTED"
			If rsTS("noreas") = 2 Then no2 = "SELECTED"
			If rsTS("noreas") = 3 Then no3 = "SELECTED"
			If rsTS("noreas") = 7 Then no7 = "SELECTED"
			If rsTS("noreas") = 8 Then no8 = "SELECTED"
			hid_q = rsTS("happen")
			hid_noreas = rsTS("noreas")
			DSnoreas = Z_CDate(rsTS("DSnoreas"))
			canbill = ""
			myStat2 = ""
			if rsTS("status") = 4 Then 
				canbill = "DISABLED"
				myStat2 = "READONLY"
			end If
			If rsTS("appDate") = cdate(sundate) Then
				If cdate(rsTS("appdate")) < cdate("2/21/2016") Then
					nooption = "<select class='seltxt' style='width: 200px;' name='noreasq" & ctr & "' " & LBcon & canbill & " '><option value='5' " & no5 & ">Appointment prior to 2/21/2016</option></select>"
				ElseIf rsTS("InstID") = 479 Then '479 2446
					nooption = "<select class='seltxt' style='width: 200px;' name='noreasq" & ctr & "' " & LBcon & canbill & " '><option value='6' " & no6 & ">No appointment hours</option></select>"
				Elseif rsTS("status") = 4 Then
					nooption = "<select class='seltxt' style='width: 200px;' name='noreasq" & ctr & "' " & LBcon & canbill & " '><option value='4' " & no4 & ">Cancelled-Billable</option></select>"
				Else
					nooption = "<select class='seltxt' style='width: 200px;' name='noreasq" & ctr & "' " & LBcon & canbill & " onchange='reqfld2();'><option value='0' " & no0 & ">&nbsp;</option><option value='1' " & no1 & ">Client did not show for the appointment</option><option value='2' " & no2 & ">Provider canceled appointment when interpreter arrived</option><option value='3' " & no3 & ">Provider had me interpret for another client not listed on vform</option><option value='7' " & no7 & ">Patient refused interpreter service</option><option value='8' " & no8 & ">No phone number provider on vform</option></select>"
				End If
				sunTS = sunTS & "<tr bgcolor='#F5F5F5'><td><input type='hidden' name='ctr" & ctr & "' value='" & rsTS("index") & "'></td><td align='center'><nobr>" & myAct & "</td>" & _
					"<td align='center'><input type='hidden' name='hid_sunq1" & ctr & "' value='" & hid_q & "'><select class='seltxt' style='width: 50px;' name='sunq1" & ctr & "' " & LBcon & canbill & " onchange='reqfld();'><option value='0' " & q0 & ">&nbsp;</option><option value='1' " & q1 & ">NO</option><option value='2' " & q2 & ">YES</option></select></td>" & _
					"<td align='center'><input type='hidden' name='hid_noreas" & ctr & "' value='" & hid_noreas & "'>" & nooption & _
					"<br>Date:<input class='main' size='11' maxlength='10' name='DSnoreas" & ctr & "' value='" & DSnoreas & "' " & LBcon & " onKeyUp=""javascript:return maskMe(this.value,this,'2,5','/');""></td>" & _
					"<td align='center'><input class='main' size='6' " & myStat2 & " maxlength='5' name='sunstart" & ctr & "' value='" & AStime & "' onKeyUp=""javascript:return maskMe(this.value,this,'2',':');"" onBlur=""javascript:return maskMe(this.value,this,'2,6',':');"" " & LBcon & "></td>" & _
					"<td align='center'><input class='main' size='6' " & myStat2 & " maxlength='5' name='sunend" & ctr & "' value='" & AEtime & "' onKeyUp=""javascript:return maskMe(this.value,this,'2',':');"" onBlur=""javascript:return maskMe(this.value,this,'2,6',':');"" " & LBcon & "></td>" & _
					"<td align='center'><input class='main' size='6' maxlength='11' readonly  name='totalhrs" & ctr & "' value='" & trim(rsTS("totalhrs")) & "'></td>" & _
					"<td align='center'>$<input class='main' size='6' " & myStat & " maxlength='5' name='suntoll" & ctr & "' value='" & Z_CZero(rsTS("toll")) & "' " & LBconToll & "></td>" & _
					"<td align='center'><input type='checkbox' disabled " &  tmpCon & "></td>" & _
					"<td align='center'><input type='checkbox' disabled " &  tmpConToll & "></td>" & _
					"<td align='center'><a href='#' onclick='uploadform(" & rsTS("index") & ");'><img src='images/upload.png' border='0' title='upload forms'></a>"  & _
					"</tr>"
				if rsTS("status") <> 4 and lbcon = "" Then
					If rsTS("deptID") = 2246 Then
						strjs = strjs & "document.frmTS.sunstart" & ctr & ".readOnly = true;" & vbcrlf & _
							"document.frmTS.sunend" & ctr & ".readOnly = true;" & vbcrlf
					Else
						strjs = strjs & "if (document.frmTS.sunq1" & ctr & ".value == 0) { " & vbcrlf & _
							"document.frmTS.sunstart" & ctr & ".value = '';" & vbcrlf & _
							"document.frmTS.sunstart" & ctr & ".readOnly = true;" & vbcrlf & _
							"document.frmTS.sunend" & ctr & ".value = '';" & vbcrlf & _
							"document.frmTS.sunend" & ctr & ".readOnly = true;" & vbcrlf & _
							"document.frmTS.noreasq"& ctr & ".disabled = true;" & vbcrlf & _
							"document.frmTS.noreasq"& ctr & ".value = 0;" & vbcrlf & _
							"document.frmTS.DSnoreas"& ctr & ".readOnly = true;" & vbcrlf & _
							"document.frmTS.DSnoreas"& ctr & ".value = '';" & vbcrlf & _
							"}" & vbcrlf & _
							"else if (document.frmTS.sunq1" & ctr & ".value == 1) { " & vbcrlf & _
							"document.frmTS.noreasq"& ctr & ".disabled = false;" & vbcrlf & _
							"document.frmTS.DSnoreas"& ctr & ".readOnly = true;" & vbcrlf & _
							"document.frmTS.sunstart" & ctr & ".readOnly = false;" & vbcrlf & _
							"document.frmTS.sunend" & ctr & ".readOnly = false;" & vbcrlf & _
							"}" & vbcrlf & _
							"else if (document.frmTS.sunq1" & ctr & ".value == 2) { " & vbcrlf & _
							"document.frmTS.sunstart" & ctr & ".readOnly = false;" & vbcrlf & _
							"document.frmTS.sunend" & ctr & ".readOnly = false;" & vbcrlf & _
							"document.frmTS.noreasq"& ctr & ".disabled = true;" & vbcrlf & _
							"document.frmTS.noreasq"& ctr & ".value = 0;" & vbcrlf & _
							"document.frmTS.DSnoreas"& ctr & ".readOnly = true;" & vbcrlf & _
							"document.frmTS.DSnoreas"& ctr & ".value = '';" & vbcrlf & _
							"}"
					End If
					strjs2 = strjs2 & "if (document.frmTS.noreasq" & ctr & ".value == 0) { " & vbcrlf & _
						"document.frmTS.DSnoreas"& ctr & ".readOnly = true;" & vbcrlf & _
						"document.frmTS.DSnoreas"& ctr & ".value = '';" & vbcrlf & _
						"}" & vbcrlf & _
						"else if (document.frmTS.noreasq" & ctr & ".value == 1) { " & vbcrlf & _
						"alert(""If the reminder call was completed, please input the date it was completed."");" & vbcrlf & _
						"document.frmTS.DSnoreas"& ctr & ".readOnly = false;" & vbcrlf & _
						"}" & vbcrlf & _
						"else if (document.frmTS.noreasq" & ctr & ".value == 2) { " & vbcrlf & _
						"document.frmTS.DSnoreas"& ctr & ".value = '';" & vbcrlf & _
						"}" & vbcrlf & _
						"else if (document.frmTS.noreasq" & ctr & ".value == 3) { " & vbcrlf & _
						"alert(""Please contact Language Bank ASAP."");" & vbcrlf & _
						"document.frmTS.DSnoreas"& ctr & ".value = '';" & vbcrlf & _
						"}"
					strjs3 = strjs3 & "if (document.frmTS.sunq1" & ctr & ".value == 1) { " & vbcrlf & _
						"if (document.frmTS.noreasq"& ctr & ".value == 0) { " & vbcrlf & _
						"alert(""Please select reason.""); " & vbcrlf & _
						"return;" & vbcrlf & _ 
						"}}"
				ElseIf rsTS("status") = 4 Then
					strjs = strjs & "document.frmTS.DSnoreas"& ctr & ".readOnly = true;" & vbcrlf & _
						"document.frmTS.DSnoreas"& ctr & ".value = '';"
				end if
			End If
			If rsTS("appDate") = cdate(sundate) + 1 Then
				If cdate(rsTS("appdate")) < cdate("2/21/2016") Then
					nooption = "<select class='seltxt' style='width: 200px;' name='noreasq" & ctr & "' " & LBcon & canbill & " '><option value='5' " & no5 & ">Appointment prior to 2/21/2016</option></select>"
				ElseIf rsTS("InstID") = 479 Then
					nooption = "<select class='seltxt' style='width: 200px;' name='noreasq" & ctr & "' " & LBcon & canbill & " '><option value='6' " & no6 & ">No appointment hours</option></select>"
				Elseif rsTS("status") = 4 Then
					nooption = "<select class='seltxt' style='width: 200px;' name='noreasq" & ctr & "' " & LBcon & canbill & " '><option value='4' " & no4 & ">Cancelled-Billable</option></select>"
				Else
					nooption = "<select class='seltxt' style='width: 200px;' name='noreasq" & ctr & "' " & LBcon & canbill & " onchange='reqfld2();'><option value='0' " & no0 & ">&nbsp;</option><option value='1' " & no1 & ">Client did not show for the appointment</option><option value='2' " & no2 & ">Provider canceled appointment when interpreter arrived</option><option value='3' " & no3 & ">Provider had me interpret for another client not listed on vform</option><option value='7' " & no7 & ">Patient refused interpreter service</option><option value='8' " & no8 & ">No phone number provider on vform</option></select>"
				End If
				monTS = monTS & "<tr bgcolor='#F5F5F5'><td><input type='hidden' name='ctr" & ctr & "' value='" & rsTS("index") & "'></td><td align='center'><nobr>" & myAct & "</td>" & _
					"<td align='center'><input type='hidden' name='hid_sunq1" & ctr & "' value='" & hid_q & "'><select class='seltxt' style='width: 50px;' name='sunq1" & ctr & "' " & LBcon & canbill & " onchange='reqfld();'><option value='0' " & q0 & ">&nbsp;</option><option value='1' " & q1 & ">NO</option><option value='2' " & q2 & ">YES</option></select></td>" & _
				  "<td align='center'><input type='hidden' name='hid_noreas" & ctr & "' value='" & hid_noreas & "'>" & nooption & _
				  "<br>Date:<input class='main' size='11' maxlength='10' name='DSnoreas" & ctr & "' value='" & DSnoreas & "' " & LBcon & " onKeyUp=""javascript:return maskMe(this.value,this,'2,5','/');""></td>" & _
				  "<td align='center'><input class='main' size='6' " & myStat2 & " maxlength='5' name='sunstart" & ctr & "' value='" & AStime & "' onKeyUp=""javascript:return maskMe(this.value,this,'2',':');"" onBlur=""javascript:return maskMe(this.value,this,'2,6',':');"" " & LBcon & "></td>" & _
					"<td align='center'><input class='main' size='6' " & myStat2 & " maxlength='5' name='sunend" & ctr & "' value='" & AEtime & "' onKeyUp=""javascript:return maskMe(this.value,this,'2',':');"" onBlur=""javascript:return maskMe(this.value,this,'2,6',':');"" " & LBcon & "></td>" & _
					"<td align='center'><input class='main' size='6' maxlength='11' readonly  name='totalhrs" & ctr & "' value='" & trim(rsTS("totalhrs")) & "'></td>" & _
					"<td align='center'>$<input class='main' size='6' " & myStat & " maxlength='5' name='suntoll" & ctr & "' value='" & Z_CZero(rsTS("toll")) & "' " & LBconToll & "></td>" & _
					"<td align='center'><input type='checkbox' disabled " &  tmpCon & "></td>" & _			
					"<td align='center'><input type='checkbox' disabled " &  tmpConToll & "></td>" & _	
					"<td align='center'><a href='#' onclick='uploadform(" & rsTS("index") & ");'><img src='images/upload.png' border='0' title='upload forms'></a>"  & _
					"</tr>"
				if rsTS("status") <> 4 and lbcon = "" Then   
					If rsTS("deptID") = 2246 Then
						strjs = strjs & "document.frmTS.sunstart" & ctr & ".readOnly = true;" & vbcrlf & _
							"document.frmTS.sunend" & ctr & ".readOnly = true;" & vbcrlf
					Else
						strjs = strjs & "if (document.frmTS.sunq1" & ctr & ".value == 0) { " & vbcrlf & _
							"document.frmTS.sunstart" & ctr & ".value = '';" & vbcrlf & _
							"document.frmTS.sunstart" & ctr & ".readOnly = true;" & vbcrlf & _
							"document.frmTS.sunend" & ctr & ".value = '';" & vbcrlf & _
							"document.frmTS.sunend" & ctr & ".readOnly = true;" & vbcrlf & _
							"document.frmTS.noreasq"& ctr & ".disabled = true;" & vbcrlf & _
							"document.frmTS.noreasq"& ctr & ".value = 0;" & vbcrlf & _
							"document.frmTS.DSnoreas"& ctr & ".readOnly = true;" & vbcrlf & _
							"document.frmTS.DSnoreas"& ctr & ".value = '';" & vbcrlf & _
							"}" & vbcrlf & _
							"else if (document.frmTS.sunq1" & ctr & ".value == 1) { " & vbcrlf & _
							"document.frmTS.noreasq"& ctr & ".disabled = false;" & vbcrlf & _
							"document.frmTS.DSnoreas"& ctr & ".readOnly = true;" & vbcrlf & _
							"document.frmTS.sunstart" & ctr & ".readOnly = false;" & vbcrlf & _
							"document.frmTS.sunend" & ctr & ".readOnly = false;" & vbcrlf & _
							"}" & vbcrlf & _
							"else if (document.frmTS.sunq1" & ctr & ".value == 2) { " & vbcrlf & _
							"document.frmTS.sunstart" & ctr & ".readOnly = false;" & vbcrlf & _
							"document.frmTS.sunend" & ctr & ".readOnly = false;" & vbcrlf & _
							"document.frmTS.noreasq"& ctr & ".disabled = true;" & vbcrlf & _
							"document.frmTS.noreasq"& ctr & ".value = 0;" & vbcrlf & _
							"document.frmTS.DSnoreas"& ctr & ".readOnly = true;" & vbcrlf & _
							"document.frmTS.DSnoreas"& ctr & ".value = '';" & vbcrlf & _
							"}"
					End If
					strjs2 = strjs2 & "if (document.frmTS.noreasq" & ctr & ".value == 0) { " & vbcrlf & _
						"document.frmTS.DSnoreas"& ctr & ".readOnly = true;" & vbcrlf & _
						"document.frmTS.DSnoreas"& ctr & ".value = '';" & vbcrlf & _
						"}" & vbcrlf & _
						"else if (document.frmTS.noreasq" & ctr & ".value == 1) { " & vbcrlf & _
						"alert(""If the reminder call was completed, please input the date it was completed."");" & vbcrlf & _
						"document.frmTS.DSnoreas"& ctr & ".readOnly = false;" & vbcrlf & _
						"}" & vbcrlf & _
						"else if (document.frmTS.noreasq" & ctr & ".value == 2) { " & vbcrlf & _
						"document.frmTS.DSnoreas"& ctr & ".value = '';" & vbcrlf & _
						"}" & vbcrlf & _
						"else if (document.frmTS.noreasq" & ctr & ".value == 3) { " & vbcrlf & _
						"alert(""Please contact Language Bank ASAP."");" & vbcrlf & _
						"document.frmTS.DSnoreas"& ctr & ".value = '';" & vbcrlf & _
						"}"
					strjs3 = strjs3 & "if (document.frmTS.sunq1" & ctr & ".value == 1) { " & vbcrlf & _
						"if (document.frmTS.noreasq"& ctr & ".value == 0) { " & vbcrlf & _
						"alert(""Please select reason.""); " & vbcrlf & _
						"return;" & vbcrlf & _ 
						"}}"
				ElseIf rsTS("status") = 4 Then
					strjs = strjs & "document.frmTS.DSnoreas"& ctr & ".readOnly = true;" & vbcrlf & _
						"document.frmTS.DSnoreas"& ctr & ".value = '';"
				end if
			End If
			If rsTS("appDate") = cdate(sundate) + 2 Then
				If cdate(rsTS("appdate")) < cdate("2/21/2016") Then
					nooption = "<select class='seltxt' style='width: 200px;' name='noreasq" & ctr & "' " & LBcon & canbill & " '><option value='5' " & no5 & ">Appointment prior to 2/21/2016</option></select>"
				ElseIf rsTS("InstID") = 479 Then
					nooption = "<select class='seltxt' style='width: 200px;' name='noreasq" & ctr & "' " & LBcon & canbill & " '><option value='6' " & no6 & ">No appointment hours</option></select>"
				Elseif rsTS("status") = 4 Then
					nooption = "<select class='seltxt' style='width: 200px;' name='noreasq" & ctr & "' " & LBcon & canbill & " '><option value='4' " & no4 & ">Cancelled-Billable</option></select>"
				Else
					nooption = "<select class='seltxt' style='width: 200px;' name='noreasq" & ctr & "' " & LBcon & canbill & " onchange='reqfld2();'><option value='0' " & no0 & ">&nbsp;</option><option value='1' " & no1 & ">Client did not show for the appointment</option><option value='2' " & no2 & ">Provider canceled appointment when interpreter arrived</option><option value='3' " & no3 & ">Provider had me interpret for another client not listed on vform</option><option value='7' " & no7 & ">Patient refused interpreter service</option><option value='8' " & no8 & ">No phone number provider on vform</option></select>"
				End If
				tueTS = tueTS & "<tr bgcolor='#F5F5F5'><td><input type='hidden' name='ctr" & ctr & "' value='" & rsTS("index") & "'></td><td align='center'><nobr>" & myAct & "</td>" & _
				 	"<td align='center'><input type='hidden' name='hid_sunq1" & ctr & "' value='" & hid_q & "'><select class='seltxt' style='width: 50px;' name='sunq1" & ctr & "' " & LBcon & canbill & " onchange='reqfld();'><option value='0' " & q0 & ">&nbsp;</option><option value='1' " & q1 & ">NO</option><option value='2' " & q2 & ">YES</option></select></td>" & _
				 	"<td align='center'><input type='hidden' name='hid_noreas" & ctr & "' value='" & hid_noreas & "'>" & nooption & _
				 	"<br>Date:<input class='main' size='11' maxlength='10' name='DSnoreas" & ctr & "' value='" & DSnoreas & "' " & LBcon & " onKeyUp=""javascript:return maskMe(this.value,this,'2,5','/');""></td>" & _
				 	"<td align='center'><input class='main' size='6' " & myStat2 & " maxlength='5' name='sunstart" & ctr & "' value='" & AStime & "' onKeyUp=""javascript:return maskMe(this.value,this,'2',':');"" onBlur=""javascript:return maskMe(this.value,this,'2,6',':');"" " & LBcon & "></td>" & _
					"<td align='center'><input class='main' size='6' " & myStat2 & " maxlength='5' name='sunend" & ctr & "' value='" & AEtime & "' onKeyUp=""javascript:return maskMe(this.value,this,'2',':');"" onBlur=""javascript:return maskMe(this.value,this,'2,6',':');"" " & LBcon & "></td>" & _
					"<td align='center'><input class='main' size='6' maxlength='11' readonly  name='totalhrs" & ctr & "' value='" & trim(rsTS("totalhrs")) & "'></td>" & _
					"<td align='center'>$<input class='main' size='6' " & myStat & " maxlength='5' name='suntoll" & ctr & "' value='" & Z_CZero(rsTS("toll")) & "' " & LBconToll & "></td>" & _
					"<td align='center'><input type='checkbox' disabled " &  tmpCon & "></td>" & _
					"<td align='center'><input type='checkbox' disabled " &  tmpConToll & "></td>" & _
					"<td align='center'><a href='#' onclick='uploadform(" & rsTS("index") & ");'><img src='images/upload.png' border='0' title='upload forms'></a>"  & _
					"</tr>"
				if rsTS("status") <> 4 and lbcon = "" Then   
					If rsTS("deptID") = 2246 Then
						strjs = strjs & "document.frmTS.sunstart" & ctr & ".readOnly = true;" & vbcrlf & _
							"document.frmTS.sunend" & ctr & ".readOnly = true;" & vbcrlf
					Else
						strjs = strjs & "if (document.frmTS.sunq1" & ctr & ".value == 0) { " & vbcrlf & _
							"document.frmTS.sunstart" & ctr & ".value = '';" & vbcrlf & _
							"document.frmTS.sunstart" & ctr & ".readOnly = true;" & vbcrlf & _
							"document.frmTS.sunend" & ctr & ".value = '';" & vbcrlf & _
							"document.frmTS.sunend" & ctr & ".readOnly = true;" & vbcrlf & _
							"document.frmTS.noreasq"& ctr & ".disabled = true;" & vbcrlf & _
							"document.frmTS.noreasq"& ctr & ".value = 0;" & vbcrlf & _
							"document.frmTS.DSnoreas"& ctr & ".readOnly = true;" & vbcrlf & _
							"document.frmTS.DSnoreas"& ctr & ".value = '';" & vbcrlf & _
							"}" & vbcrlf & _
							"else if (document.frmTS.sunq1" & ctr & ".value == 1) { " & vbcrlf & _
							"document.frmTS.noreasq"& ctr & ".disabled = false;" & vbcrlf & _
							"document.frmTS.DSnoreas"& ctr & ".readOnly = true;" & vbcrlf & _
							"document.frmTS.sunstart" & ctr & ".readOnly = false;" & vbcrlf & _
							"document.frmTS.sunend" & ctr & ".readOnly = false;" & vbcrlf & _
							"}" & vbcrlf & _
							"else if (document.frmTS.sunq1" & ctr & ".value == 2) { " & vbcrlf & _
							"document.frmTS.sunstart" & ctr & ".readOnly = false;" & vbcrlf & _
							"document.frmTS.sunend" & ctr & ".readOnly = false;" & vbcrlf & _
							"document.frmTS.noreasq"& ctr & ".disabled = true;" & vbcrlf & _
							"document.frmTS.noreasq"& ctr & ".value = 0;" & vbcrlf & _
							"document.frmTS.DSnoreas"& ctr & ".readOnly = true;" & vbcrlf & _
							"document.frmTS.DSnoreas"& ctr & ".value = '';" & vbcrlf & _
							"}"
					End If
					strjs2 = strjs2 & "if (document.frmTS.noreasq" & ctr & ".value == 0) { " & vbcrlf & _
						"document.frmTS.DSnoreas"& ctr & ".readOnly = true;" & vbcrlf & _
						"document.frmTS.DSnoreas"& ctr & ".value = '';" & vbcrlf & _
						"}" & vbcrlf & _
						"else if (document.frmTS.noreasq" & ctr & ".value == 1) { " & vbcrlf & _
						"alert(""If the reminder call was completed, please input the date it was completed."");" & vbcrlf & _
						"document.frmTS.DSnoreas"& ctr & ".readOnly = false;" & vbcrlf & _
						"}" & vbcrlf & _
						"else if (document.frmTS.noreasq" & ctr & ".value == 2) { " & vbcrlf & _
						"document.frmTS.DSnoreas"& ctr & ".value = '';" & vbcrlf & _
						"}" & vbcrlf & _
						"else if (document.frmTS.noreasq" & ctr & ".value == 3) { " & vbcrlf & _
						"alert(""Please contact Language Bank ASAP."");" & vbcrlf & _
						"document.frmTS.DSnoreas"& ctr & ".value = '';" & vbcrlf & _
						"}"
					strjs3 = strjs3 & "if (document.frmTS.sunq1" & ctr & ".value == 1) { " & vbcrlf & _
						"if (document.frmTS.noreasq"& ctr & ".value == 0) { " & vbcrlf & _
						"alert(""Please select reason.""); " & vbcrlf & _
						"return;" & vbcrlf & _ 
						"}}"
				ElseIf rsTS("status") = 4 Then
					strjs = strjs & "document.frmTS.DSnoreas"& ctr & ".readOnly = true;" & vbcrlf & _
						"document.frmTS.DSnoreas"& ctr & ".value = '';"
				end if
			End If
			If rsTS("appDate") = cdate(sundate) + 3 Then
				If cdate(rsTS("appdate")) < cdate("2/21/2016") Then
					nooption = "<select class='seltxt' style='width: 200px;' name='noreasq" & ctr & "' " & LBcon & canbill & " '><option value='5' " & no5 & ">Appointment prior to 2/21/2016</option></select>"
				ElseIf rsTS("InstID") = 479 Then
					nooption = "<select class='seltxt' style='width: 200px;' name='noreasq" & ctr & "' " & LBcon & canbill & " '><option value='6' " & no6 & ">No appointment hours</option></select>"
				Elseif rsTS("status") = 4 Then
					nooption = "<select class='seltxt' style='width: 200px;' name='noreasq" & ctr & "' " & LBcon & canbill & " '><option value='4' " & no4 & ">Cancelled-Billable</option></select>"
				Else
					nooption = "<select class='seltxt' style='width: 200px;' name='noreasq" & ctr & "' " & LBcon & canbill & " onchange='reqfld2();'><option value='0' " & no0 & ">&nbsp;</option><option value='1' " & no1 & ">Client did not show for the appointment</option><option value='2' " & no2 & ">Provider canceled appointment when interpreter arrived</option><option value='3' " & no3 & ">Provider had me interpret for another client not listed on vform</option><option value='7' " & no7 & ">Patient refused interpreter service</option><option value='8' " & no8 & ">No phone number provider on vform</option></select>"
				End If
				wedTS = wedTS & "<tr bgcolor='#F5F5F5'><td><input type='hidden' name='ctr" & ctr & "' value='" & rsTS("index") & "'></td><td align='center'><nobr>" & myAct & "</td>" & _
				 	"<td align='center'><input type='hidden' name='hid_sunq1" & ctr & "' value='" & hid_q & "'><select class='seltxt' style='width: 50px;' name='sunq1" & ctr & "' " & LBcon & canbill & " onchange='reqfld();'><option value='0' " & q0 & ">&nbsp;</option><option value='1' " & q1 & ">NO</option><option value='2' " & q2 & ">YES</option></select></td>" & _
				 	"<td align='center'><input type='hidden' name='hid_noreas" & ctr & "' value='" & hid_noreas & "'>" & nooption & _
				 	"<br>Date:<input class='main' size='11' maxlength='10' name='DSnoreas" & ctr & "' value='" & DSnoreas & "' " & LBcon & " onKeyUp=""javascript:return maskMe(this.value,this,'2,5','/');""></td>" & _
				 	"<td align='center'><input class='main' size='6' " & myStat2 & " maxlength='5' name='sunstart" & ctr & "' value='" & AStime & "' onKeyUp=""javascript:return maskMe(this.value,this,'2',':');"" onBlur=""javascript:return maskMe(this.value,this,'2,6',':');"" " & LBcon & "></td>" & _
					"<td align='center'><input class='main' size='6' " & myStat2 & " maxlength='5' name='sunend" & ctr & "' value='" & AEtime & "' onKeyUp=""javascript:return maskMe(this.value,this,'2',':');"" onBlur=""javascript:return maskMe(this.value,this,'2,6',':');"" " & LBcon & "></td>" & _
					"<td align='center'><input class='main' size='6' maxlength='11' readonly  name='totalhrs" & ctr & "' value='" & trim(rsTS("totalhrs")) & "'></td>" & _
					"<td align='center'>$<input class='main' size='6' " & myStat & " maxlength='5' name='suntoll" & ctr & "' value='" & Z_CZero(rsTS("toll")) & "' " & LBconToll & "></td>" & _
					"<td align='center'><input type='checkbox' disabled " &  tmpCon & "></td>" & _
					"<td align='center'><input type='checkbox' disabled " &  tmpConToll & "></td>" & _
					"<td align='center'><a href='#' onclick='uploadform(" & rsTS("index") & ");'><img src='images/upload.png' border='0' title='upload forms'></a>"  & _
					"</tr>"
				if rsTS("status") <> 4 and lbcon = "" Then   
					If rsTS("deptID") = 2246 Then
						strjs = strjs & "document.frmTS.sunstart" & ctr & ".readOnly = true;" & vbcrlf & _
							"document.frmTS.sunend" & ctr & ".readOnly = true;" & vbcrlf
					Else
						strjs = strjs & "if (document.frmTS.sunq1" & ctr & ".value == 0) { " & vbcrlf & _
							"document.frmTS.sunstart" & ctr & ".value = '';" & vbcrlf & _
							"document.frmTS.sunstart" & ctr & ".readOnly = true;" & vbcrlf & _
							"document.frmTS.sunend" & ctr & ".value = '';" & vbcrlf & _
							"document.frmTS.sunend" & ctr & ".readOnly = true;" & vbcrlf & _
							"document.frmTS.noreasq"& ctr & ".disabled = true;" & vbcrlf & _
							"document.frmTS.noreasq"& ctr & ".value = 0;" & vbcrlf & _
							"document.frmTS.DSnoreas"& ctr & ".readOnly = true;" & vbcrlf & _
							"document.frmTS.DSnoreas"& ctr & ".value = '';" & vbcrlf & _
							"}" & vbcrlf & _
							"else if (document.frmTS.sunq1" & ctr & ".value == 1) { " & vbcrlf & _
							"document.frmTS.noreasq"& ctr & ".disabled = false;" & vbcrlf & _
							"document.frmTS.DSnoreas"& ctr & ".readOnly = true;" & vbcrlf & _
							"document.frmTS.sunstart" & ctr & ".readOnly = false;" & vbcrlf & _
							"document.frmTS.sunend" & ctr & ".readOnly = false;" & vbcrlf & _
							"}" & vbcrlf & _
							"else if (document.frmTS.sunq1" & ctr & ".value == 2) { " & vbcrlf & _
							"document.frmTS.sunstart" & ctr & ".readOnly = false;" & vbcrlf & _
							"document.frmTS.sunend" & ctr & ".readOnly = false;" & vbcrlf & _
							"document.frmTS.noreasq"& ctr & ".disabled = true;" & vbcrlf & _
							"document.frmTS.noreasq"& ctr & ".value = 0;" & vbcrlf & _
							"document.frmTS.DSnoreas"& ctr & ".readOnly = true;" & vbcrlf & _
							"document.frmTS.DSnoreas"& ctr & ".value = '';" & vbcrlf & _
							"}"
					End If
					strjs2 = strjs2 & "if (document.frmTS.noreasq" & ctr & ".value == 0) { " & vbcrlf & _
						"document.frmTS.DSnoreas"& ctr & ".readOnly = true;" & vbcrlf & _
						"document.frmTS.DSnoreas"& ctr & ".value = '';" & vbcrlf & _
						"}" & vbcrlf & _
						"else if (document.frmTS.noreasq" & ctr & ".value == 1) { " & vbcrlf & _
						"alert(""If the reminder call was completed, please input the date it was completed."");" & vbcrlf & _
						"document.frmTS.DSnoreas"& ctr & ".readOnly = false;" & vbcrlf & _
						"}" & vbcrlf & _
						"else if (document.frmTS.noreasq" & ctr & ".value == 2) { " & vbcrlf & _
						"document.frmTS.DSnoreas"& ctr & ".value = '';" & vbcrlf & _
						"}" & vbcrlf & _
						"else if (document.frmTS.noreasq" & ctr & ".value == 3) { " & vbcrlf & _
						"alert(""Please contact Language Bank ASAP."");" & vbcrlf & _
						"document.frmTS.DSnoreas"& ctr & ".value = '';" & vbcrlf & _
						"}"
					strjs3 = strjs3 & "if (document.frmTS.sunq1" & ctr & ".value == 1) { " & vbcrlf & _
						"if (document.frmTS.noreasq"& ctr & ".value == 0) { " & vbcrlf & _
						"alert(""Please select reason.""); " & vbcrlf & _
						"return;" & vbcrlf & _ 
						"}}"
				ElseIf rsTS("status") = 4 Then
					strjs = strjs & "document.frmTS.DSnoreas"& ctr & ".readOnly = true;" & vbcrlf & _
						"document.frmTS.DSnoreas"& ctr & ".value = '';"
				end if
			End If
			If rsTS("appDate") = cdate(sundate) + 4 Then
				If cdate(rsTS("appdate")) < cdate("2/21/2016") Then
					nooption = "<select class='seltxt' style='width: 200px;' name='noreasq" & ctr & "' " & LBcon & canbill & " '><option value='5' " & no5 & ">Appointment prior to 2/21/2016</option></select>"
				ElseIf rsTS("InstID") = 479 Then
					nooption = "<select class='seltxt' style='width: 200px;' name='noreasq" & ctr & "' " & LBcon & canbill & " '><option value='6' " & no6 & ">No appointment hours</option></select>"
				Elseif rsTS("status") = 4 Then
					nooption = "<select class='seltxt' style='width: 200px;' name='noreasq" & ctr & "' " & LBcon & canbill & " '><option value='4' " & no4 & ">Cancelled-Billable</option></select>"
				Else
					nooption = "<select class='seltxt' style='width: 200px;' name='noreasq" & ctr & "' " & LBcon & canbill & " onchange='reqfld2();'><option value='0' " & no0 & ">&nbsp;</option><option value='1' " & no1 & ">Client did not show for the appointment</option><option value='2' " & no2 & ">Provider canceled appointment when interpreter arrived</option><option value='3' " & no3 & ">Provider had me interpret for another client not listed on vform</option><option value='7' " & no7 & ">Patient refused interpreter service</option><option value='8' " & no8 & ">No phone number provider on vform</option></select>"
				End If
				thuTS = thuTS & "<tr bgcolor='#F5F5F5'><td><input type='hidden' name='ctr" & ctr & "' value='" & rsTS("index") & "'></td><td align='center'><nobr>" & myAct & "</td>" & _
				 	"<td align='center'><input type='hidden' name='hid_sunq1" & ctr & "' value='" & hid_q & "'><select class='seltxt' style='width: 50px;' name='sunq1" & ctr & "' " & LBcon & canbill & " onchange='reqfld();'><option value='0' " & q0 & ">&nbsp;</option><option value='1' " & q1 & ">NO</option><option value='2' " & q2 & ">YES</option></select></td>" & _
				 	"<td align='center'><input type='hidden' name='hid_noreas" & ctr & "' value='" & hid_noreas & "'>" & nooption & _
				 	"<br>Date:<input class='main' size='11' maxlength='10' name='DSnoreas" & ctr & "' value='" & DSnoreas & "' " & LBcon & " onKeyUp=""javascript:return maskMe(this.value,this,'2,5','/');""></td>" & _
				 	"<td align='center'><input class='main' size='6' " & myStat2 & " maxlength='5' name='sunstart" & ctr & "' value='" & AStime & "' onKeyUp=""javascript:return maskMe(this.value,this,'2',':');"" onBlur=""javascript:return maskMe(this.value,this,'2,6',':');"" " & LBcon & "></td>" & _
					"<td align='center'><input class='main' size='6' " & myStat2 & " maxlength='5' name='sunend" & ctr & "' value='" & AEtime & "' onKeyUp=""javascript:return maskMe(this.value,this,'2',':');"" onBlur=""javascript:return maskMe(this.value,this,'2,6',':');"" " & LBcon & "></td>" & _
					"<td align='center'><input class='main'size='6' maxlength='11' readonly  name='totalhrs" & ctr & "' value='" & trim(rsTS("totalhrs")) & "'></td>" & _
					"<td align='center'>$<input class='main' size='6' " & myStat & " maxlength='5' name='suntoll" & ctr & "' value='" & Z_CZero(rsTS("toll")) & "' " & LBconToll & "></td>" & _
					"<td align='center'><input type='checkbox' disabled " &  tmpCon & "></td>" & _
					"<td align='center'><input type='checkbox' disabled " &  tmpConToll & "></td>" & _
					"<td align='center'><a href='#' onclick='uploadform(" & rsTS("index") & ");'><img src='images/upload.png' border='0' title='upload forms'></a>"  & _
					"</tr>"
				if rsTS("status") <> 4 and lbcon = "" Then   
					If rsTS("deptID") = 2246 Then
						strjs = strjs & "document.frmTS.sunstart" & ctr & ".readOnly = true;" & vbcrlf & _
							"document.frmTS.sunend" & ctr & ".readOnly = true;" & vbcrlf
					Else
						strjs = strjs & "if (document.frmTS.sunq1" & ctr & ".value == 0) { " & vbcrlf & _
							"document.frmTS.sunstart" & ctr & ".value = '';" & vbcrlf & _
							"document.frmTS.sunstart" & ctr & ".readOnly = true;" & vbcrlf & _
							"document.frmTS.sunend" & ctr & ".value = '';" & vbcrlf & _
							"document.frmTS.sunend" & ctr & ".readOnly = true;" & vbcrlf & _
							"document.frmTS.noreasq"& ctr & ".disabled = true;" & vbcrlf & _
							"document.frmTS.noreasq"& ctr & ".value = 0;" & vbcrlf & _
							"document.frmTS.DSnoreas"& ctr & ".readOnly = true;" & vbcrlf & _
							"document.frmTS.DSnoreas"& ctr & ".value = '';" & vbcrlf & _
							"}" & vbcrlf & _
							"else if (document.frmTS.sunq1" & ctr & ".value == 1) { " & vbcrlf & _
							"document.frmTS.noreasq"& ctr & ".disabled = false;" & vbcrlf & _
							"document.frmTS.DSnoreas"& ctr & ".readOnly = true;" & vbcrlf & _
							"document.frmTS.sunstart" & ctr & ".readOnly = false;" & vbcrlf & _
							"document.frmTS.sunend" & ctr & ".readOnly = false;" & vbcrlf & _
							"}" & vbcrlf & _
							"else if (document.frmTS.sunq1" & ctr & ".value == 2) { " & vbcrlf & _
							"document.frmTS.sunstart" & ctr & ".readOnly = false;" & vbcrlf & _
							"document.frmTS.sunend" & ctr & ".readOnly = false;" & vbcrlf & _
							"document.frmTS.noreasq"& ctr & ".disabled = true;" & vbcrlf & _
							"document.frmTS.noreasq"& ctr & ".value = 0;" & vbcrlf & _
							"document.frmTS.DSnoreas"& ctr & ".readOnly = true;" & vbcrlf & _
							"document.frmTS.DSnoreas"& ctr & ".value = '';" & vbcrlf & _
							"}"
					End If
					strjs2 = strjs2 & "if (document.frmTS.noreasq" & ctr & ".value == 0) { " & vbcrlf & _
						"document.frmTS.DSnoreas"& ctr & ".readOnly = true;" & vbcrlf & _
						"document.frmTS.DSnoreas"& ctr & ".value = '';" & vbcrlf & _
						"}" & vbcrlf & _
						"else if (document.frmTS.noreasq" & ctr & ".value == 1) { " & vbcrlf & _
						"alert(""If the reminder call was completed, please input the date it was completed."");" & vbcrlf & _
						"document.frmTS.DSnoreas"& ctr & ".readOnly = false;" & vbcrlf & _
						"}" & vbcrlf & _
						"else if (document.frmTS.noreasq" & ctr & ".value == 2) { " & vbcrlf & _
						"document.frmTS.DSnoreas"& ctr & ".value = '';" & vbcrlf & _
						"}" & vbcrlf & _
						"else if (document.frmTS.noreasq" & ctr & ".value == 3) { " & vbcrlf & _
						"alert(""Please contact Language Bank ASAP."");" & vbcrlf & _
						"document.frmTS.DSnoreas"& ctr & ".value = '';" & vbcrlf & _
						"}"
					strjs3 = strjs3 & "if (document.frmTS.sunq1" & ctr & ".value == 1) { " & vbcrlf & _
						"if (document.frmTS.noreasq"& ctr & ".value == 0) { " & vbcrlf & _
						"alert(""Please select reason.""); " & vbcrlf & _
						"return;" & vbcrlf & _ 
						"}}"
				ElseIf rsTS("status") = 4 Then
					strjs = strjs & "document.frmTS.DSnoreas"& ctr & ".readOnly = true;" & vbcrlf & _
						"document.frmTS.DSnoreas"& ctr & ".value = '';"
				end if
			End If
			If rsTS("appDate") = cdate(sundate) + 5 Then
				If cdate(rsTS("appdate")) < cdate("2/21/2016") Then
					nooption = "<select class='seltxt' style='width: 200px;' name='noreasq" & ctr & "' " & LBcon & canbill & " '><option value='5' " & no5 & ">Appointment prior to 2/21/2016</option></select>"
				ElseIf rsTS("InstID") = 479 Then
					nooption = "<select class='seltxt' style='width: 200px;' name='noreasq" & ctr & "' " & LBcon & canbill & " '><option value='6' " & no6 & ">No appointment hours</option></select>"
				Elseif rsTS("status") = 4 Then
					nooption = "<select class='seltxt' style='width: 200px;' name='noreasq" & ctr & "' " & LBcon & canbill & " '><option value='4' " & no4 & ">Cancelled-Billable</option></select>"
				Else
					nooption = "<select class='seltxt' style='width: 200px;' name='noreasq" & ctr & "' " & LBcon & canbill & " onchange='reqfld2();'><option value='0' " & no0 & ">&nbsp;</option><option value='1' " & no1 & ">Client did not show for the appointment</option><option value='2' " & no2 & ">Provider canceled appointment when interpreter arrived</option><option value='3' " & no3 & ">Provider had me interpret for another client not listed on vform</option><option value='7' " & no7 & ">Patient refused interpreter service</option><option value='8' " & no8 & ">No phone number provider on vform</option></select>"
				End If
				friTS = friTS & "<tr bgcolor='#F5F5F5'><td><input type='hidden' name='ctr" & ctr & "' value='" & rsTS("index") & "'></td><td align='center'><nobr>" & myAct & "</td>" & _
				 	"<td align='center'><input type='hidden' name='hid_sunq1" & ctr & "' value='" & hid_q & "'><select class='seltxt' style='width: 50px;' name='sunq1" & ctr & "' " & LBcon & canbill & " onchange='reqfld();'><option value='0' " & q0 & ">&nbsp;</option><option value='1' " & q1 & ">NO</option><option value='2' " & q2 & ">YES</option></select></td>" & _
				 	"<td align='center'><input type='hidden' name='hid_noreas" & ctr & "' value='" & hid_noreas & "'>" & nooption & _
				 	"<br>Date:<input class='main' size='11' maxlength='10' name='DSnoreas" & ctr & "' value='" & DSnoreas & "' " & LBcon & " onKeyUp=""javascript:return maskMe(this.value,this,'2,5','/');""></td>" & _
				 	"<td align='center'><input class='main' size='6' " & myStat2 & " maxlength='5' name='sunstart" & ctr & "' value='" & AStime & "' onKeyUp=""javascript:return maskMe(this.value,this,'2',':');"" onBlur=""javascript:return maskMe(this.value,this,'2,6',':');"" " & LBcon & "></td>" & _
					"<td align='center'><input class='main' size='6' " & myStat2 & " maxlength='5' name='sunend" & ctr & "' value='" & AEtime & "' onKeyUp=""javascript:return maskMe(this.value,this,'2',':');"" onBlur=""javascript:return maskMe(this.value,this,'2,6',':');"" " & LBcon & "></td>" & _
					"<td align='center'><input class='main' size='6' maxlength='11' readonly  name='totalhrs" & ctr & "' value='" & trim(rsTS("totalhrs")) & "'></td>" & _
					"<td align='center'>$<input class='main' size='6' " & myStat & " maxlength='5' name='suntoll" & ctr & "' value='" & Z_CZero(rsTS("toll")) & "' " & LBconToll & "></td>" & _
					"<td align='center'><input type='checkbox' disabled " &  tmpCon & "></td>" & _
					"<td align='center'><input type='checkbox' disabled " &  tmpConToll & "></td>" & _
					"<td align='center'><a href='#' onclick='uploadform(" & rsTS("index") & ");'><img src='images/upload.png' border='0' title='upload forms'></a>"  & _
					"</tr>"
				if rsTS("status") <> 4 and lbcon = "" Then   
					If rsTS("deptID") = 2246 Then
						strjs = strjs & "document.frmTS.sunstart" & ctr & ".readOnly = true;" & vbcrlf & _
							"document.frmTS.sunend" & ctr & ".readOnly = true;" & vbcrlf
					Else
						strjs = strjs & "if (document.frmTS.sunq1" & ctr & ".value == 0) { " & vbcrlf & _
							"document.frmTS.sunstart" & ctr & ".value = '';" & vbcrlf & _
							"document.frmTS.sunstart" & ctr & ".readOnly = true;" & vbcrlf & _
							"document.frmTS.sunend" & ctr & ".value = '';" & vbcrlf & _
							"document.frmTS.sunend" & ctr & ".readOnly = true;" & vbcrlf & _
							"document.frmTS.noreasq"& ctr & ".disabled = true;" & vbcrlf & _
							"document.frmTS.noreasq"& ctr & ".value = 0;" & vbcrlf & _
							"document.frmTS.DSnoreas"& ctr & ".readOnly = true;" & vbcrlf & _
							"document.frmTS.DSnoreas"& ctr & ".value = '';" & vbcrlf & _
							"}" & vbcrlf & _
							"else if (document.frmTS.sunq1" & ctr & ".value == 1) { " & vbcrlf & _
							"document.frmTS.noreasq"& ctr & ".disabled = false;" & vbcrlf & _
							"document.frmTS.DSnoreas"& ctr & ".readOnly = true;" & vbcrlf & _
							"document.frmTS.sunstart" & ctr & ".readOnly = false;" & vbcrlf & _
							"document.frmTS.sunend" & ctr & ".readOnly = false;" & vbcrlf & _
							"}" & vbcrlf & _
							"else if (document.frmTS.sunq1" & ctr & ".value == 2) { " & vbcrlf & _
							"document.frmTS.sunstart" & ctr & ".readOnly = false;" & vbcrlf & _
							"document.frmTS.sunend" & ctr & ".readOnly = false;" & vbcrlf & _
							"document.frmTS.noreasq"& ctr & ".disabled = true;" & vbcrlf & _
							"document.frmTS.noreasq"& ctr & ".value = 0;" & vbcrlf & _
							"document.frmTS.DSnoreas"& ctr & ".readOnly = true;" & vbcrlf & _
							"document.frmTS.DSnoreas"& ctr & ".value = '';" & vbcrlf & _
							"}"
					End If
					strjs2 = strjs2 & "if (document.frmTS.noreasq" & ctr & ".value == 0) { " & vbcrlf & _
						"document.frmTS.DSnoreas"& ctr & ".readOnly = true;" & vbcrlf & _
						"document.frmTS.DSnoreas"& ctr & ".value = '';" & vbcrlf & _
						"}" & vbcrlf & _
						"else if (document.frmTS.noreasq" & ctr & ".value == 1) { " & vbcrlf & _
						"alert(""If the reminder call was completed, please input the date it was completed."");" & vbcrlf & _
						"document.frmTS.DSnoreas"& ctr & ".readOnly = false;" & vbcrlf & _
						"}" & vbcrlf & _
						"else if (document.frmTS.noreasq" & ctr & ".value == 2) { " & vbcrlf & _
						"document.frmTS.DSnoreas"& ctr & ".value = '';" & vbcrlf & _
						"}" & vbcrlf & _
						"else if (document.frmTS.noreasq" & ctr & ".value == 3) { " & vbcrlf & _
						"alert(""Please contact Language Bank ASAP."");" & vbcrlf & _
						"document.frmTS.DSnoreas"& ctr & ".value = '';" & vbcrlf & _
						"}"
					strjs3 = strjs3 & "if (document.frmTS.sunq1" & ctr & ".value == 1) { " & vbcrlf & _
						"if (document.frmTS.noreasq"& ctr & ".value == 0) { " & vbcrlf & _
						"alert(""Please select reason.""); " & vbcrlf & _
						"return;" & vbcrlf & _ 
						"}}"
				ElseIf rsTS("status") = 4 Then
					strjs = strjs & "document.frmTS.DSnoreas"& ctr & ".readOnly = true;" & vbcrlf & _
						"document.frmTS.DSnoreas"& ctr & ".value = '';"
				end if
			End If
			If rsTS("appDate") = cdate(satdate) Then
				If cdate(rsTS("appdate")) < cdate("2/21/2016") Then
					nooption = "<select class='seltxt' style='width: 200px;' name='noreasq" & ctr & "' " & LBcon & canbill & " '><option value='5' " & no5 & ">Appointment prior to 2/21/2016</option></select>"
				ElseIf rsTS("InstID") = 479 Then
					nooption = "<select class='seltxt' style='width: 200px;' name='noreasq" & ctr & "' " & LBcon & canbill & " '><option value='6' " & no6 & ">No appointment hours</option></select>"
				Elseif rsTS("status") = 4 Then
					nooption = "<select class='seltxt' style='width: 200px;' name='noreasq" & ctr & "' " & LBcon & canbill & " '><option value='4' " & no4 & ">Cancelled-Billable</option></select>"
				Else
					nooption = "<select class='seltxt' style='width: 200px;' name='noreasq" & ctr & "' " & LBcon & canbill & " onchange='reqfld2();'><option value='0' " & no0 & ">&nbsp;</option><option value='1' " & no1 & ">Client did not show for the appointment</option><option value='2' " & no2 & ">Provider canceled appointment when interpreter arrived</option><option value='3' " & no3 & ">Provider had me interpret for another client not listed on vform</option><option value='7' " & no7 & ">Patient refused interpreter service</option><option value='8' " & no8 & ">No phone number provider on vform</option></select>"
				End If
				satTS = satTS & "<tr bgcolor='#F5F5F5'><td><input type='hidden' name='ctr" & ctr & "' value='" & rsTS("index") & "'></td><td align='center'><nobr>" & myAct & "</td>" & _
				 	"<td align='center'><input type='hidden' name='hid_sunq1" & ctr & "' value='" & hid_q & "'><select class='seltxt' style='width: 50px;' name='sunq1" & ctr & "' " & LBcon & canbill & " onchange='reqfld();'><option value='0' " & q0 & ">&nbsp;</option><option value='1' " & q1 & ">NO</option><option value='2' " & q2 & ">YES</option></select></td>" & _
				 	"<td align='center'><input type='hidden' name='hid_noreas" & ctr & "' value='" & hid_noreas & "'>" & nooption & _
				 	"<br>Date:<input class='main' size='11' maxlength='10' name='DSnoreas" & ctr & "' value='" & DSnoreas & "' " & LBcon & " onKeyUp=""javascript:return maskMe(this.value,this,'2,5','/');""></td>" & _
				 	"<td align='center'><input class='main' size='6' " & myStat2 & " maxlength='5' name='sunstart" & ctr & "' value='" & AStime & "' onKeyUp=""javascript:return maskMe(this.value,this,'2',':');"" onBlur=""javascript:return maskMe(this.value,this,'2,6',':');"" " & LBcon & "></td>" & _
					"<td align='center'><input class='main' size='6' " & myStat2 & " maxlength='5' name='sunend" & ctr & "' value='" & AEtime & "' onKeyUp=""javascript:return maskMe(this.value,this,'2',':');"" onBlur=""javascript:return maskMe(this.value,this,'2,6',':');"" " & LBcon & "></td>" & _
					"<td align='center'><input class='main' size='6' maxlength='11' readonly name='totalhrs" & ctr & "' value='" & trim(rsTS("totalhrs")) & "'></td>" & _
					"<td align='center'>$<input class='main' size='6' " & myStat & " maxlength='5' name='suntoll" & ctr & "' value='" & Z_CZero(rsTS("toll")) & "' " & LBconToll & "></td>" & _
					"<td align='center'><input type='checkbox' disabled " &  tmpCon & "></td>" & _
					"<td align='center'><input type='checkbox' disabled " &  tmpConToll & "></td>" & _
					"<td align='center'><a href='#' onclick='uploadform(" & rsTS("index") & ");'><img src='images/upload.png' border='0' title='upload forms'></a>"  & _
					"</tr>"
				if rsTS("status") <> 4 and lbcon = "" Then   
					If rsTS("deptID") = 2246 Then
						strjs = strjs & "document.frmTS.sunstart" & ctr & ".readOnly = true;" & vbcrlf & _
							"document.frmTS.sunend" & ctr & ".readOnly = true;" & vbcrlf
					Else
						strjs = strjs & "if (document.frmTS.sunq1" & ctr & ".value == 0) { " & vbcrlf & _
							"document.frmTS.sunstart" & ctr & ".value = '';" & vbcrlf & _
							"document.frmTS.sunstart" & ctr & ".readOnly = true;" & vbcrlf & _
							"document.frmTS.sunend" & ctr & ".value = '';" & vbcrlf & _
							"document.frmTS.sunend" & ctr & ".readOnly = true;" & vbcrlf & _
							"document.frmTS.noreasq"& ctr & ".disabled = true;" & vbcrlf & _
							"document.frmTS.noreasq"& ctr & ".value = 0;" & vbcrlf & _
							"document.frmTS.DSnoreas"& ctr & ".readOnly = true;" & vbcrlf & _
							"document.frmTS.DSnoreas"& ctr & ".value = '';" & vbcrlf & _
							"}" & vbcrlf & _
							"else if (document.frmTS.sunq1" & ctr & ".value == 1) { " & vbcrlf & _
							"document.frmTS.noreasq"& ctr & ".disabled = false;" & vbcrlf & _
							"document.frmTS.DSnoreas"& ctr & ".readOnly = true;" & vbcrlf & _
							"document.frmTS.sunstart" & ctr & ".readOnly = false;" & vbcrlf & _
							"document.frmTS.sunend" & ctr & ".readOnly = false;" & vbcrlf & _
							"}" & vbcrlf & _
							"else if (document.frmTS.sunq1" & ctr & ".value == 2) { " & vbcrlf & _
							"document.frmTS.sunstart" & ctr & ".readOnly = false;" & vbcrlf & _
							"document.frmTS.sunend" & ctr & ".readOnly = false;" & vbcrlf & _
							"document.frmTS.noreasq"& ctr & ".disabled = true;" & vbcrlf & _
							"document.frmTS.noreasq"& ctr & ".value = 0;" & vbcrlf & _
							"document.frmTS.DSnoreas"& ctr & ".readOnly = true;" & vbcrlf & _
							"document.frmTS.DSnoreas"& ctr & ".value = '';" & vbcrlf & _
							"}"
					End If
					strjs2 = strjs2 & "if (document.frmTS.noreasq" & ctr & ".value == 0) { " & vbcrlf & _
						"document.frmTS.DSnoreas"& ctr & ".readOnly = true;" & vbcrlf & _
						"document.frmTS.DSnoreas"& ctr & ".value = '';" & vbcrlf & _
						"}" & vbcrlf & _
						"else if (document.frmTS.noreasq" & ctr & ".value == 1) { " & vbcrlf & _
						"alert(""If the reminder call was completed, please input the date it was completed."");" & vbcrlf & _
						"document.frmTS.DSnoreas"& ctr & ".readOnly = false;" & vbcrlf & _
						"}" & vbcrlf & _
						"else if (document.frmTS.noreasq" & ctr & ".value == 2) { " & vbcrlf & _
						"document.frmTS.DSnoreas"& ctr & ".value = '';" & vbcrlf & _
						"}" & vbcrlf & _
						"else if (document.frmTS.noreasq" & ctr & ".value == 3) { " & vbcrlf & _
						"alert(""Please contact Language Bank ASAP."");" & vbcrlf & _
						"document.frmTS.DSnoreas"& ctr & ".value = '';" & vbcrlf & _
						"}"
					strjs3 = strjs3 & "if (document.frmTS.sunq1" & ctr & ".value == 1) { " & vbcrlf & _
						"if (document.frmTS.noreasq"& ctr & ".value == 0) { " & vbcrlf & _
						"alert(""Please select reason.""); " & vbcrlf & _
						"return;" & vbcrlf & _ 
						"}}"
				ElseIf rsTS("status") = 4 Then
					strjs = strjs & "document.frmTS.DSnoreas"& ctr & ".readOnly = true;" & vbcrlf & _
						"document.frmTS.DSnoreas"& ctr & ".value = '';"
				end if
			End If
			ctr = ctr + 1
			rsTS.MoveNext
		Loop
	End If
	rsTS.Close
	Set rsClose = Nothing
	if ctr = 0 then disave = "disabled"
%>
<html>
	<head>
		<title>Language Bank - Interpreter appointments</title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function uploadform(xxx)
		{
			newwindow = window.open('upload.asp?reqid=' + xxx,'name','height=175,width=400,scrollbars=0,directories=0,status=0,toolbar=0,resizable=0');
			if (window.focus) {newwindow.focus()}
		}
		function reqfld2() {
			<%=strjs2%>
		}
		function reqfld() {
			<%=strjs%>
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
		function SaveTS()
		{
			<%=strjs3%>
			var ans = window.confirm("Save Entries?\n*Must email or drop off or fax your receipt to Languagebank.");
			if (ans){
				document.frmTS.action = "tsheetaction.asp?action=1";
				document.frmTS.submit();
			}
			document.frmTS.action = "tsheetaction.asp?action=1";
			document.frmTS.submit();
	
		}
		function CalendarView(strDate)
		{
			document.frmTS.action = 'tsheet.asp?tmpdate=' + strDate;
			document.frmTS.submit();
		}
		function PrevMonth()
		{
			document.frmTS.action = "tsheet.asp?action=1&tmpDate=" + '<%=sundate%>';
			document.frmTS.submit();	
		}
		function NextMonth()
		{
			document.frmTS.action = "tsheet.asp?action=2&tmpDate=" + '<%=sundate%>';
			document.frmTS.submit();	
		}
		//-->
		</script>
	</head>
	<body onload="reqfld();">
		<table cellSpacing='0' cellPadding='0' height='100%' width="100%" border='0' class='bgstyle2'>
					<tr>
						<td valign='top'>
							<!-- #include file="_header.asp" -->
						</td>
					</tr>
					<tr>
						<td valign='top' >
							<table cellSpacing='0' cellPadding='0' width="100%" border='0'>
									<!-- #include file="_greetme.asp" -->
								<tr>
								<td class='title' colspan='10' align='center'><nobr> Interpreter appointments</td>
								</tr>
								<tr>
									<td  align='center' colspan='12'>
										<div name="dErr" style="width: 250px; height:55px;OVERFLOW: auto;">
											<table border='0' cellspacing='1'>		
												<tr>
													<td><span class='error'><%=Session("MSG")%></span></td>
												</tr>
											</table>
										</div>
									</td>
								</tr>
								<tr>
									<td align='right' width='150px'>Name:</td>
									<td class='confirm'><%=GetIntr(Session("UIntr"))%></td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td align='right'>Date:</td>
									<td class='confirm'><%=sundate%> - <%=satdate%></td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									
									<td colspan='2' align='center'>
										<form name='frmTS' method='POST'>
											<table border='0' cellpadding='1' cellspacing='2' width='90%'>
												<tr>
													<td align='center' class='tblgrn'>Date</td>
													<td align='center' class='tblgrn'>Activity</td>
													<td align='center' class='tblgrn'>Did you interpret for the name on the verification form</td>
													<td align='center' class='tblgrn'>Reason<br>(for NO answers)</td>
													<td align='center' class='tblgrn'>Appt. Start Time</td>
													<td align='center' class='tblgrn'>Appt. End Time</td>
													<td align='center' class='tblgrn'>Total Hours</td>
													<td align='center' class='tblgrn'>Tolls & parking<br>with receipts</td>
													<td align='center' class='tblgrn'>Approved Time</td>
													<td align='center' class='tblgrn'>Approved Tolls & parking</td>
													<td align='center' class='tblgrn'>Upload documents</td>
												</tr>	
												<tr><td align='center' class='confirm'>SUN</td></tr>
												<%=sunTS%>
												<tr><td align='center' class='confirm'>MON</td></tr>
												<%=monTS%>
												<tr><td align='center' class='confirm'>TUE</td></tr>
												<%=tueTS%>
												<tr><td align='center' class='confirm'>WED</td></tr>
												<%=wedTS%>
												<tr><td align='center' class='confirm'>THU</td></tr>
												<%=thuTS%>
												<tr><td align='center' class='confirm'>FRI</td></tr>
												<%=friTS%>
												<tr><td align='center' class='confirm'>SAT</td></tr>
												<%=satTS%>
											</table>
										
									</td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td colspan='12' align='center'>
										*To <b>SAVE</b>, answer question first then input Appt. Start time,App. End time and Tolls & parking then click on 'Save' button.<br>
										PLEASE USE MILITARY TIME (24-Hour FORMAT). Do not use 00:00 / midnight on both fields. If you need to enter it, use 00:01.<br>
										Once approved, you can no longer edit Appt. Start Time,Appt. End Time and Tolls & parking
									</td>
								</tr>
								<tr>
									<td colspan='12' align='center' height='100px' valign='bottom'>
										<input type='hidden' name='tmpDate' value="<%=sundate%>">
										<input type='hidden' name='tmpDate2' value="<%=satdate%>">
										<input type='hidden' name='myCTR' value="<%=ctr%>">
										<input type='hidden' name='myCTR2' value="<%=ctrCon%>">
										<input class='btn' type='button' value='<Prev Week'  onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='PrevMonth();'>
										<input class='btn' type='button' value='Save Entries' <%=billedna%> <%=disave%> onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='SaveTS();'>
										<input class='btn' type='button' value='Next Week>'  onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='NextMonth();'>
										
									</td>
								</tr>
								</form>
							</table>
						</td>
					</tr>
					<tr>
						<td valign='bottom'>
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