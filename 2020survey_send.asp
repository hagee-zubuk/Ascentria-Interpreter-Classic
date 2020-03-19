<!doctype html>
<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%
Function Z_YMDDate(dtDate)
DIM lngTmp, strDay, strTmp
	If Not IsDate(dtDate) Then Exit Function
	strTmp = DatePart("yyyy", dtDate)
	Z_YMDDate = strTmp & "-"

	lngTmp = CLng(DatePart("m", dtDate))
	If lngTmp < 10 Then Z_YMDDate = Z_YMDDate & "0"
	Z_YMDDate = Z_YMDDate & lngTmp & "-"
	lngTmp = CLng(DatePart("d", dtDate))
	If lngTmp < 10 Then Z_YMDDate = Z_YMDDate & "0"
	Z_YMDDate = Z_YMDDate & lngTmp
End Function

' just save the thing 
IntrID = Session("UIntr")
'strSQL = "SELECT COUNT([id]) AS c FROM [2020Survey] WHERE [IntrID]=" & IntrID & " AND [SvyDt]='" & Z_YMDDate(Date) & "'"
Set rsZ = Server.CreateObject("ADODB.Recordset")
rsZ.Open "[2020Survey]", g_strCONNDB, 1, 3
rsZ.AddNew
rsZ("ts") = Now
rsZ("SvyDt") = Date
rsZ("IntrID") = IntrID
rsZ("Q1") = Request("Q1")
rsZ("Q2") = Request("Q2")
rsZ("Q3") = Request("Q3")
rsZ("Sig") = Request("txtSig")
rsZ.Update
rsZ.Close

Set rsZ = Nothing
Response.Redirect "calendarview2.asp"

%>
