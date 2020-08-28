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
chk_01	= Request("chk_01")
chk_02	= Request("chk_02")
chk_03	= Request("chk_03")
fname	= Request("fname")
mname	= Request("mname")
lname	= Request("lname")
suffix	= Request("suffix")
strIP	= Request.ServerVariables("REMOTE_ADDR")
strUA	= Request.ServerVariables("HTTP_USER_AGENT")
strCook	= Request.ServerVariables("HTTP_COOKIE")
strHost	= Request.ServerVariables("HTTP_REFERER")
strRef	= Request.ServerVariables("HTTP_HOST")


If (chk_01<>"1") Or (chk_02<>"1") Or (chk_03<>"1") Or (chkSig<>"1") Then
	Response.Write "Yer fucked<br />"
	Response.Write "CHK 01: [" & chk_01 & "]<br />"
	Response.Write "CHK 02: [" & chk_02 & "]<br />"
	Response.Write "CHK 03: [" & chk_03 & "]<br />"
	Response.Write "CHK Sg: [" & chkSig & "]<br />"
	Response.End
	Session("MSG") = "Please check all the boxes and type your name."
	Response.Redirect("msd_ia.asp")
End If
strSQL = "SELECT * FROM [Emp_MSD] WHERE [userid]=" & userid
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
rsIR("chk_01")	= chk_01
rsIR("chk_02")	= chk_02
rsIR("chk_03")	= chk_03
rsIr("last") 	= Now()
rsIR("referrer")	= strRef
rsIR("useragent")	= strUA
rsIR.Update
rsIR.Close
Set rsIR = Nothing

Response.Redirect "msd_done.asp?fetch=" & userid
%>
' ese_proc.asp