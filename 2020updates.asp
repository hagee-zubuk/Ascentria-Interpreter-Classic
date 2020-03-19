<!doctype html>
<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%
Dim strUPLFolder, strFiles, dtLastMod, tmpLastMod, tmpFilePath, latestFile
strUPLFolder = uploadpath & "00.Updates\"
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(strUPLFolder)
Set colFiles = objFolder.Files
strFiles 	= ""
tmpFile 	= ""
dtLastMod = CDate("2020-01-01")
For Each objFile in colFiles
	strFiles = strFiles & "<tr><td>" & objFile.Name & "</td>"
	tmpLastMod = CDate(objFile.DateLastModified)
	strFiles = strFiles & "<td>" & objFile.DateLastModified & "</td></tr>" & vbCrLf
	If (tmpLastMod > dtLastMod) Then
		tmpFile = objFile.Path
		tmpType = Z_GetExt(objFile.Name)
		dtLastMod = tmpLastMod
		latestFile = objFile.Name
	End If
Next

Dim strContent, objFileToRead
blnDisp = False
If (tmpType = "TXT") Then
	blnDisp = True
	Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile(tmpFile,1)
	strContent = objFileToRead.ReadAll()
	objFileToRead.Close
	Set objFileToRead = Nothing
	strContent = "<pre>" & vbCrLf & strContent & vbCrLf & "</pre>"
ElseIf (tmpType = "HTML") OR (tmpType = "HTM") Then
	blnDisp = True
	Dim strLine, blnRd
	strContent = "<div class=""htmlcontent"">" & vbCrLf & vbCrLf
	blnRd = False
	Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile(tmpFile,1)
	Do While Not objFileToRead.AtEndOfStream
		strLine = objFileToRead.ReadLine()
		If blnRd Then
			strContent = strContent & vbTab & strLine & vbCrLf
		End If
		If InStr(0, strLine, "<body>") > 0 Then
			blnRd = True
		ElseIf InStr(0, strLine, "</body>") > 0 Then
			blnRd = False
		End If
	Loop
	objFileToRead.Close
	Set objFileToRead = Nothing
	strContent = strContent & vbCrLf & vbCrLf & "</div>"
End If

If Not blnDisp Then ' download the update file
	latestFile = Replace(latestFile, " ", "" )
	' force a download
	Set oFileStream = Server.CreateObject("ADODB.Stream")
	oFileStream.Open
	oFileStream.Type = 1 'Binary
On Error Resume Next
	oFileStream.LoadFromFile tmpFile
	If Err.Number <> 0 Then
		Response.Write "Error: " & Err.Number & " [" & Err.Message & "]"
		Response.End
	End If
	Response.Clear
	If (tmpType = "PDF") Then
		Response.ContentType = "application/pdf"
		Response.AddHeader "Content-Disposition", "attachment; filename=" & latestFile
	Else
		Response.ContentType = "application/octet-stream"
		Response.AddHeader "Content-Disposition", "attachment; filename=v-update." & tmpType
	End If

	Dim lSize, lBlocks
	'Const CHUNK = 2048
	Const CHUNK = 20480
	lSize = oFileStream.Size
	Response.AddHeader "Content-Size", lSize
	lBlocks = 1
	Response.Buffer = False
	Do Until oFileStream.EOS Or Not Response.IsClientConnected
		Response.BinaryWrite(oFileStream.Read(CHUNK))
	Loop
	'Response.BinaryWrite oFileStream.Read
	oFileStream.Close
	Set oFileStream= Nothing
	Response.End
End If
%>
<html lang="en">
<head>
	<meta charset="utf-8" />
	<meta name="viewport" content="width=device-width, initial-scale=1.0" />
	<meta name="description" content="COVID-19 Update page" />
	<meta name="author" content="Zubuk PH" />

	<title>Language Bank - COVID-19 Update</title>
	<link type="text/css" rel="stylesheet" href="style.css" />
	<link type="text/css" rel="stylesheet" href="CalendarControl.css" />
	<script src="CalendarControl.js" language="javascript"></script>
	<style>
div.container { margin: 0px 20px; }
pre { font-size: 130%; border: 1px dotted gray; padding: 10px; }
	</style>
</head>
<body>
<table cellSpacing='0' cellPadding='0' height='100%' width="100%" border='0' class='bgstyle2'>
	<tr><td valign='top'><!-- #include file="_header.asp" --></td></tr>
</table>
<div class="container">
<h1>COVID-19 Update</h1>
File date: <%=dtLastMod%><br />
<%= strContent %>
<p>Latest file: <b><%=latestFile %></b> (<%= tmpFile%> ) <%= tmpType %></p>
<!-- <p>Previous files:</p>
<table>
	<tr><th>Name</th><th>Modified</th></tr>
<%= strFiles %>
</table>
--></div>
<table cellSpacing='0' cellPadding='0' width="100%" border='0' class='bgstyle2' style="position: absolute; bottom: 0px;">
	<tr><td valign="bottom"><!-- #include file="_footer.asp" --></td></tr>
</table>
</body>
</html>