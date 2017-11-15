<%
'paths needed
DIM 	g_strCONN, g_strDBPath

g_strCONNDB = "Provider=SQLOLEDB;Data Source=ERNIE\SQLEXPRESS;Initial Catalog=langbank;Integrated Security=SSPI;"
Set g_strCONN = Server.CreateObject("ADODB.Connection")
g_strCONN.Open g_strCONNDB

'HIST SQL
g_strCONNDB2 = "Provider=SQLOLEDB;Data Source=ERNIE\SQLEXPRESS;Initial Catalog=histLB;Integrated Security=SSPI;"
Set g_strCONNHIST2 = Server.CreateObject("ADODB.Connection")
g_strCONNHIST2.Open g_strCONNDB2

RepPath = "C:\WORK\ascentria\interpreter\CSV\"
RepPath2 = "C:\WORK\ascentria\interpreter\CSV\"
RepCSV = "/CSV/"
RepCSV2 = "/CSV/"
BackupStr = "C:\WORK\ascentria\interpreter\CSV\"
pdfStr = "C:\WORK\ascentria\interpreter\PDF\"
EmailLog = "C:\WORK\ascentria\interpreter\log\EmailLog.txt"
LoginLog = "C:\WORK\ascentria\interpreter\log\LoginLog.txt"
AdminLog = "C:\WORK\ascentria\interpreter\log\AdminLog.txt"

'HistoryDB = "C:\work\LSS-LBIS\db\HistLangBank.mdb"
'g_strCONNHist = "PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" & HistoryDB & ";"

'FOR HOSPITALPILOT
'g_strDBPathHP = "C:\work\InterReq\db\interpreter.mdb"
'g_strCONNHP = "PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" & g_strDBPathHP & ";"
g_strCONNHPs = "Provider=SQLOLEDB;Data Source=ERNIE\SQLEXPRESS;Initial Catalog=interpreterSQL;Integrated Security=SSPI;"
Set g_strCONNHP = Server.CreateObject("ADODB.Connection")
g_strCONNHP.Open g_strCONNHPs

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
g_strCONNDBupload = "Provider=SQLOLEDB;Data Source=ERNIE\SQLEXPRESS;Initial Catalog=langbankuploads;Integrated Security=SSPI;"
Set g_strCONNupload = Server.CreateObject("ADODB.Connection")
g_strCONNupload.Open g_strCONNDBupload

googlemapskey = "AIzaSyAHcSoJYxk465hDVj1_wMXTAozARDkfFgo"
SurveyPath = "C:\WORK\ascentria\interpreter\DHHSsurvey\"
DirectionPath = "C:\WORK\ascentria\interpreter\misc\"
%>