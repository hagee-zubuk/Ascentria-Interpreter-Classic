<%

'redirect user to default page when not logged in
If Request("PDF") <> 1 Then
	tmpUser = Request.Cookies("LBUSER")
	If tmpUser = "" Then
		Session("MSG") = "ERROR: Cookies has expired or was not found.<br> Please sign in again."
		Response.redirect "default.asp"
	End If
	tmpName	= Request.Cookies("LBUsrName")
	If tmpName = "" Then
		Session("MSG") = "ERROR: Session has expired.<br> Please sign in again."
		Response.redirect "default.asp"
	End If
	If Session("UIntr") = "" Then
		Session("MSG") = "ERROR: Session has expired.<br> Please sign in again."
		Response.redirect "default.asp"
	End If
	If Session("UsrID") = "" Then
		Session("MSG") = "ERROR: Session has expired.<br> Please sign in again."
		Response.redirect "default.asp"
	End If
End If

Function ZZ_YMDDate(dtDate)
DIM lngTmp, strDay, strTmp
	If Not IsDate(dtDate) Then Exit Function
	strTmp = DatePart("yyyy", dtDate)
	ZZ_YMDDate = strTmp & "-"

	lngTmp = CLng(DatePart("m", dtDate))
	If lngTmp < 10 Then ZZ_YMDDate = ZZ_YMDDate & "0"
	ZZ_YMDDate = ZZ_YMDDate & lngTmp & "-"
	lngTmp = CLng(DatePart("d", dtDate))
	If lngTmp < 10 Then ZZ_YMDDate = ZZ_YMDDate & "0"
	ZZ_YMDDate = ZZ_YMDDate & lngTmp
End Function


'check if interpreter has filled something up for today
'ZIntrID = CLng(Session("UIntr"))
'If (ZIntrID > 0) Then
''	strSQL = "SELECT COUNT([id]) AS c FROM [2020Survey] WHERE [IntrID]=" & ZIntrID & " AND [SvyDt]='" & ZZ_YMDDate(Date) & "'"
''	Set rsZ = Server.CreateObject("ADODB.Recordset")
''	rsZ.Open strSQL, g_strCONNDB, 3, 1
''	If rsZ("c") = 0 Then
''		rsZ.Close
''		Set rsZ = Nothing
''		Response.Redirect "2020survey.asp"
''	End If
''	rsZ.Close
''	Set rsZ = Nothing
'End If
%>