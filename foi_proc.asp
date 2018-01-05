<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
userid	= Request("userid")
empname	= Request("empname")
addr	= Request("addr")
cellno	= Request("cellno")
email	= Request("email")
chksig	= Request("chkSig")
fname	= Request("fname")
mname	= Request("mname")
lname	= Request("lname")
suffix	= Request("suffix")
strIP	= Request.ServerVariables("REMOTE_ADDR")
strUA	= Request.ServerVariables("HTTP_USER_AGENT")
strCook	= Request.ServerVariables("HTTP_COOKIE")
strHost	= Request.ServerVariables("HTTP_REFERER")
strRef	= Request.ServerVariables("HTTP_HOST")

strSQL = "SELECT * FROM [InfoRelease] WHERE [userid]=" & userid
Set rsIR = Server.CreateObject("ADODB.RecordSet")
rsIR.Open strSQL, g_strCONN, 1, 3
If rsIR.EOF Then
	rsIR.AddNew
	rsIr("userid") = userid
	rsIr("ts") = Now()
End If
rsIR("empname") = empname
rsIR("addr")	= addr
rsIR("cellno")	= cellno
rsIR("email")	= email
rsIR("chksig")	= chkSig
rsIR("fname")	= fname
rsIR("mname")	= mname
rsIR("lname")	= lname
rsIR("suffix")	= suffix
rsIR("ip")		= strIP
rsIR("cookies")	= strCook
rsIR("host")	= strhost
rsIr("last") 	= Now()
rsIR("referrer")	= strRef
rsIR("useragent")	= strUA
rsIR.Update
rsIR.Close
Set rsIR = Nothing

Response.Redirect "foi_done.asp?fetch=" & userid
%>