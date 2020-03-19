<%Language=VBScript%>
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<%Response.AddHeader "Pragma", "No-Cache" %>
<%
ValidAko = False
ChangePass = False

Function IsOnCall(xxx)
	IsOnCall = 0
	Set rsOC = Server.CreateObject("ADODB.RecordSet")
	sqlOC = "SELECT oncall FROM Interpreter_T WHERE [index] = " & xxx
	rsOC.Open sqlOC, g_strCONN, 3, 1
	If Not rsOC.EOF Then
		If rsOC("oncall") Then IsOnCall = 1
	End If
	rsOC.Close
	Set rsOC = Nothing
End Function
Function Z_SendOnce(intrID)
	If Z_CZero(intrID) > 0 Then
		Set rsIntr = Server.CreateObject("ADODB.RecordSet")
		rsIntr.Open "SELECT sendonce FROM Interpreter_T WHERE [index] = " & intrID, g_strCONN, 1, 3
		If Not rsIntr.EOF Then
			rsIntr("sendonce") = False
			rsIntr.Update
		End If
		rsIntr.Close
		Set rsIntr = Nothing
	End If
End Function
Set rsUser = Server.CreateObject("ADODB.RecordSet")
sqlUser = "SELECT * FROM User_t WHERE upper(Username) = '" & ucase(Request("txtUN")) & "' "
rsUser.Open sqlUser, g_strCONN, 3, 1
If Not rsUser.EOF Then
	Response.Cookies("LBUSER") = Request("txtUN")
	If Request("txtPW") = Z_DoDecrypt(rsUser("password")) Then 
		If rsUser("type") = 2 Then
			Response.Cookies("LBUSERTYPE") = rsUser("type") 'gets user type - admin/default
			If rsUser("type") = 2 Then 
				Session("UIntr") = rsUser("IntrID")
				Call Z_SendOnce(rsUser("IntrID"))
			End If
			Response.Cookies("ONCALL") = IsOnCall(Session("UIntr"))
			Session("UsrName") = rsUser("Fname") & " " & rsUser("Lname")
			Response.Cookies("LBUsrName") = rsUser("Fname") & " " & rsUser("Lname")
			Session("UsrID") = rsUser("index")
			If rsUser("Lname") = "" Then 
				Session("UsrName") = Session("UsrName") & " " & rsUser("Lname")
				Response.Cookies("LBUsrName") = Session("UsrName") & " " & rsUser("Lname")
			End If
			If rsUser("reset") Then ChangePass = True
			ValidAko = True
		Else
			Session("MSG") = "ERROR: Invalid user type. Only Interpreters are allowed in this site."
		End If
	Else
		Session("MSG") = "ERROR: Invalid username and/or password."
	End If
Else
	Session("MSG") = "ERROR: Invalid username and/or password."
End If
rsUser.Close
Set rsUser = Nothing
<!-- #include file="_closeSQL.asp" -->
If ValidAko = True Then
	'CREATE LOG
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set LogMe = fso.OpenTextFile(LoginLog, 8, True)
	strLog = Now & vbtab & "Successful Sign in :: User: " & Session("UsrName") & " -- DMZ1"
	LogMe.WriteLine strLog
	Set LogMe = Nothing
	Set fso = Nothing
	If Request.Cookies("LBUSERTYPE") <> 2 Then
		Response.Redirect "calendarview2.asp"
	Else
		If ChangePass Then
			Response.Redirect "chngpass.asp"
		Else	
			Response.Redirect "2020survey.asp" 'calendarview2.asp"
		End If
	End If
Else
	'CREATE LOG
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set LogMe = fso.OpenTextFile(LoginLog, 8, True)
	strLog = Now & vbtab & "Error in Sign in :: User: " & Request("txtUN") & " -- DMZ1"
	LogMe.WriteLine strLog
	Set LogMe = Nothing
	Set fso = Nothing
	Response.Redirect "default.asp"
End If
%>