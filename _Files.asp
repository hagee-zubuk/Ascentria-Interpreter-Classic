<%
'paths needed
DIM 	g_strCONN, g_strDBPath

'g_strDBPath = "C:\work\LSS-LBIS\db\LangBank.mdb"
'g_strCONN = "PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" & g_strDBPath & ";"10.10.1.35
g_strCONNDB = "Provider=SQLOLEDB;Data Source=10.10.16.35;Initial Catalog=langbank;Integrated Security=SSPI;"
'g_strCONNDB = "Provider=SQLOLEDB;Data Source=10.10.1.35;Initial Catalog=langbank;Integrated Security=SSPI;"
Set g_strCONN = Server.CreateObject("ADODB.Connection")
g_strCONN.Open g_strCONNDB

'HIST SQL
g_strCONNDB2 = "Provider=SQLOLEDB;Data Source=10.10.16.35;Initial Catalog=histLB;Integrated Security=SSPI;"
'g_strCONNDB2 = "Provider=SQLOLEDB;Data Source=10.10.1.35;Initial Catalog=histLB;Integrated Security=SSPI;"
Set g_strCONNHIST2 = Server.CreateObject("ADODB.Connection")
g_strCONNHIST2.Open g_strCONNDB2

RepPath = "C:\work\LSS-LBIS\web\CSV\"
RepPath2 = "C:\work\LSS-LBIS\web\CSV\"
RepCSV = "/CSV/"
RepCSV2 = "/CSV/"
BackupStr = "C:\work\LSS-LBIS\CSV\"
pdfStr = "C:\work\LSS-LBIS\PDF\"
EmailLog = "c:\work\lss-lbis\log\EmailLog.txt"
LoginLog = "c:\work\lss-lbis\log\LoginLog.txt"
AdminLog = "c:\work\lss-lbis\log\AdminLog.txt"

'HistoryDB = "C:\work\LSS-LBIS\db\HistLangBank.mdb"
'g_strCONNHist = "PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" & HistoryDB & ";"

'FOR HOSPITALPILOT
'g_strDBPathHP = "C:\work\InterReq\db\interpreter.mdb"
'g_strCONNHP = "PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" & g_strDBPathHP & ";"
g_strCONNHPs = "Provider=SQLOLEDB;Data Source=10.10.16.35;Initial Catalog=interpreterSQL;Integrated Security=SSPI;"
'Set g_strCONNHP = Server.CreateObject("ADODB.Connection")
'g_strCONNHP.Open g_strCONNHPs
g_strCONNHP = g_strCONNHPs 

'HIST SQL
'g_strCONNDB2 = "Provider=SQLOLEDB;Data Source=10.10.16.35;Initial Catalog=histLB;Integrated Security=SSPI;"
'Set g_strCONNHIST2 = Server.CreateObject("ADODB.Connection")
'g_strCONNHIST2.Open g_strCONNDB2

'FOR WIZARD DB
g_strDBPathW = "C:\work\LSS-LBIS\db\LBWizard.mdb"
g_strCONNW = "PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" & g_strDBPathW & ";"

'FOR INTERPRETER TRACKING
'g_strCONNDB3 = "Provider=SQLOLEDB;Data Source=10.10.16.35;Initial Catalog=langbankappt;Integrated Security=SSPI;"
'Set g_strCONNIntr = Server.CreateObject("ADODB.Connection")
'g_strCONNIntr.Open g_strCONNDB3

'upload path
uploadpath = "\\10.10.16.35\Interpreter_Upload\"

'FOR UPLOAD
g_strCONNDBupload = "Provider=SQLOLEDB;Data Source=10.10.16.35;Initial Catalog=langbankuploads;Integrated Security=SSPI;"
Set g_strCONNupload = Server.CreateObject("ADODB.Connection")
g_strCONNupload.Open g_strCONNDBupload

googlemapskey = "AIzaSyAHcSoJYxk465hDVj1_wMXTAozARDkfFgo"
SurveyPath = "C:\work\LSS-LBIS\DHHSsurvey\"
DirectionPath = "C:\work\LSS-LBIS\misc\"
%>
<!-- #include file="_zEmail.asp" -->