<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%
Function mkArray(n) ' Small utility function
	Dim s
	s = Space(n)
	mkArray = Split(s," ")
End Function

lngID = Session("UIntr")
If lngID < 1 Then
	Session("MSG") = "survey response index is missing"
	Response.Redirect "survey.v18.asp"
End If

strSrc = Request("URL")
arrUrl = Split(strSrc, "/")

strServerName = Request("SERVER_NAME")
strNm = "Survey" & lngID & ".18.pdf"
strPDF = pdfStr & strNm 
If Request("HTTPS") = "on" Then
	strUrl = "https://"
Else
	strUrl = "http://"
End If
strUrl = strUrl & strServerName & "/"
For lngI = 1 To UBound(arrUrl) - 1
	strUrl = strUrl & arrUrl(lngI) & "/"
Next
strUrl = strUrl & "survey.report.asp?ix=" & lngID

Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(strPDF) Then
	Call fso.DeleteFile(strPDF, TRUE)
End If

Set rsSurv = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT [release] FROM [surveyreports] WHERE [iid]=" & lngID
rsSurv.Open strSQL, g_strCONN, 3, 1
blnRelease = FALSE
If Not rsSurv.EOF Then
	blnRelease = CBool( rsSurv("release") )
End If
rsSurv.Close

'blnRelease = True
If Not blnRelease Then
	Response.Redirect "survey.v18.asp"
End If

'On Error Resume Next
'Set rsSurv = Server.CreateObject("ADODB.RecordSet")
Set theDoc = Server.CreateObject("ABCpdf9.Doc")
thedoc.HtmlOptions.PageCacheClear
theDoc.HtmlOptions.RetryCount = 3
theDoc.HtmlOptions.Timeout = 120000
theDoc.Pos.X = 10
theDoc.Pos.Y = 10
theDoc.Rect.Inset 50, 50
theDoc.Page = theDoc.AddPage()

theID = theDoc.AddImageUrl(strUrl, True, 1200, True)

Do
  theDoc.Framerect
  If Not theDoc.Chainable(theID) Then Exit Do
  theDoc.Page = theDoc.AddPage()
  theID = theDoc.AddImageToChain(theID)
Loop

For i = 1 to theDoc.PageCount
     theDoc.PageNumber = i
     theDoc.Flatten
Next

'If Err.Number <> 0 Then
	Err.Clear
	strSQL = "SELECT [release], [viewed] FROM [surveyreports] WHERE [iid]=" & lngID
	rsSurv.Open strSQL, g_strCONN, 1, 3
	If Not rsSurv.EOF Then
		rsSurv("viewed") = Now
		rsSurv.Update
	End If
	rsSurv.Close

'theDoc.Save "C:\work\apr_pdf\zz.pdf"
theData = theDoc.GetData() 

theDoc.Save strPDF
theDoc.Clear
Response.Addheader "Content-Disposition", "attachment; filename=""InSurvey.pdf"""
Response.AddHeader "content-length", UBound(theData) - LBound(theData) + 1 
Response.Addheader "Content-Transfer-Encoding", "binary"
Response.ExpiresAbsolute = #January 1, 2001 01:00:00#
Response.CacheControl = "Private"
Response.ContentType = "application/pdf"

Response.BinaryWrite theData
Response.Flush
Response.End


''		Response.Clear
''		'Response.Status = "206 Partial Content"
''		Response.Addheader "Content-Disposition", "attachment; filename=""InSurvey.pdf"""
''		Response.Addheader "Content-Length", intFSz 
''		Response.Addheader "Accept-Ranges", "bytes"
''
''		Set BinaryStream = CreateObject("ADODB.Stream")
''		BinaryStream.Type = 1
''		BinaryStream.Open
''		BinaryStream.LoadFromFile strPDF
''		binCont = BinaryStream.Read
''		BinaryStream.Close
''		Response.BinaryWrite binCont
''		Response.Flush()
''		Set BinaryStream = Nothing

'Response.End
%>