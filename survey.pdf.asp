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

blnRelease = True
If Not blnRelease Then
	Response.Redirect "survey.v18.asp"
End If

'On Error Resume Next
'Set rsSurv = Server.CreateObject("ADODB.RecordSet")
Set theDoc = Server.CreateObject("ABCpdf9.Doc")
theDoc.HtmlOptions.PageCacheClear
theDoc.HtmlOptions.RetryCount = 3
theDoc.HtmlOptions.Timeout = 120000
theDoc.Pos.X = 10
theDoc.Pos.Y = 10
theID = theDoc.AddImageUrl(strUrl)
For i = 1 To theDoc.PageCount
	theDoc.PageNumber = i
	theDoc.Flatten
Next

theDoc.Save strPDF	
'If Err.Number <> 0 Then
	Err.Clear
	strSQL = "SELECT [release] FROM [surveyreports] WHERE [iid]=" & lngID
	rsSurv.Open strSQL, g_strCONN, 1, 3
	If Not rsSurv.EOF Then
		rsSurv("viewed") = Now
		rsSurv.Update
	End If
	rsSurv.Close

	If fso.FileExists(strPDF) Then
		Set objFile = fso.GetFile(strPDF)
		intFSz = objFile.Size
		Set objFile = Nothing
'Response.Write "<h1>ready</h1><p>to download</p>"
'Response.End
		Response.Clear
		'Response.Status = "206 Partial Content"
		Response.Addheader "Content-Disposition", "attachment; filename=""InSurvey.pdf"""
		Response.Addheader "Content-Length", intFSz 
		Response.Addheader "Accept-Ranges", "bytes"
		Response.Addheader "Content-Transfer-Encoding", "binary"
		Response.ExpiresAbsolute = #January 1, 2001 01:00:00#
		Response.CacheControl = "Private"
		Response.ContentType = "application/pdf"

		Set BinaryStream = CreateObject("ADODB.Stream")
		BinaryStream.Type = 1
		BinaryStream.Open
		BinaryStream.LoadFromFile strPDF
		binCont = BinaryStream.Read
		BinaryStream.Close
		Response.BinaryWrite binCont
		Response.Flush()
		Set BinaryStream = Nothing
	Else
		Err.Clear
		Response.Clear
		'Response.Write "woah!"		
		Response.Status = "404 File Not Found"
	End If
'End If
Set rsSurv = Nothing

'Response.End
%>