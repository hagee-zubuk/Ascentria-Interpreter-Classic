<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%
Function Z_MakeUniqueFileName()
	tmpdate = replace(date, "/", "") 
	tmpTime = replace(FormatDateTime(time, 3), ":", "")
	tmpTime = replace(tmpTime, " ", "")
	Z_MakeUniqueFileName = tmpdate & tmpTime
End Function

lngRID = Z_CLng(Request("reqid"))

Set oUpload = Server.CreateObject("SCUpload.Upload")
oUpload.Upload
lngRID	= oUpload.Form("reqid")
lngType	= oUpload.Form("utype")
If oUpload.Files.Count = 0 Then
	Set oUpload = Nothing
	Response.Write "<h1>Please specify a file to import (0" & lngType & "-" & lngRID & ").</h1>"
	Response.Write "<a href=""viewuploads.asp?reqid=" & lngRID & """>try again</a>"
	Session("MSG") = ""
	Response.End
Else
	folderpath = uploadpath & lngRID
	folderpathvform =	uploadpath & lngRID & "\vform"
	folderpathtoll 	=	uploadpath & lngRID & "\tolls"
	Set fso = Server.CreateObject("Scripting.FileSystemObject")

	ServerShare = Z_UNFixPath(uploadpath)
	UserName = "acadatasrv2\LB_Webserv"
	Password = "1LBVerifyIdentity1"
	Set NetworkObject = CreateObject("WScript.Network")
On Error Resume Next
	NetworkObject.MapNetworkDrive "", ServerShare, False, UserName, Password
	If Err.Number<>0 Then
		Response.Write "Mapping: " & ServerShare & "<br />"
		Response.Write "<p>Error connecting to file repository!</p><pre>" & Err.Description & "</pre><br />Please contact LanguageBank<br />"
		Response.End
	End If
On Error Goto 0

	If Not fso.FolderExists(folderpath) Then fso.CreateFolder(folderpath)
	If Not fso.FolderExists(folderpathvform) Then fso.CreateFolder(folderpathvform)
	If Not fso.FolderExists(folderpathtoll) Then fso.CreateFolder(folderpathtoll)

	oFileName = oUpload.Files(1).Item(1).filename
	strExt = LCase(Z_GetExt(oFileName))
	UniqueFilename = Z_MakeUniqueFileName()
	If (lngType = 0) Then
		filename = "vform"
		folder = folderpathvform
	Else
		filename = "tollsandpark"
		folder = folderpathtoll
	End If
	filename = filename & UniqueFilename & "." & strExt
	oUpload.Files(1).Item(1).Save "C:\WORK\LSS-LBIS\uploads", filename
	srcF = "C:\WORK\LSS-LBIS\uploads\" & filename
	fso.CopyFile srcF, Z_FixPath(folder), True
	Session("MSG") = "File Saved."

	NetworkObject.RemoveNetworkDrive ServerShare, True, False
	Set NetworkObject = Nothing
	

	Set rsUpload = Server.CreateObject("ADODB.RecordSet")
	strSQL = "SELECT * FROM uploads"
	rsUpload.Open strSQL, g_strCONNupload, 1, 3
	rsUpload.AddNew
	rsUpload("RID") = lngRID
	rsUpload("type") = lngType
	rsUpload("filename") = FileName
	rsUpload("timestamp") = Now
	rsUpload("staff") = 0
	rsUpload.Update
	rsUpload.Close
	Set rsUpload = Nothing

	fso.DeleteFile srcF, True
	Set fso = Nothing
End If

Response.Redirect "viewuploads.asp?reqid=" & lngRID
%>