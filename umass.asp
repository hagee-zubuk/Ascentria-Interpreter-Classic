<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
Set theDoc = Server.CreateObject("ABCpdf9.Doc") 'converts html to pdf
Set theDoc2 = Server.CreateObject("ABCpdf9.Doc")
Set theDoc3 = Server.CreateObject("ABCpdf9.Doc")

fname = "VerificationForm" & Request("ReqID") & "UM.pdf"
attachPDF = pdfStr & fname
strUrl = "https://interpreter.thelanguagebank.org/interpreter/umass-body.asp?ReqID=" & Request("ReqID")
thedoc.HtmlOptions.PageCacheClear
theDoc.HtmlOptions.RetryCount = 3
theDoc.HtmlOptions.Timeout = 120000
theDoc.Pos.X = 10
theDoc.Pos.Y = 10
theID = theDoc.AddImageUrl(strUrl)
'theDoc.Page = theDoc.AddPage()
theDoc2.Read(DirectionPath & "Instructions for Interpreters at UMass pdf version 10.10.17.pdf")
theDoc.Append(theDoc2)
theDoc3.Read(DirectionPath & "umass_encounter_form.2018.pdf")
theDoc.Append(theDoc3)

For i = 1 To theDoc.PageCount
  theDoc.PageNumber = i
  theDoc.Flatten
Next
	
theDoc.Save attachPDF
	
Set theDoc2 = Nothing
Set theDoc = Nothing 
' Response.Write attachPDF
'save datestamp and save in hist
' Set rsTS = Server.CreateObject("ADODB.RecordSet")
' rsTS.Open "UPDATE request_T SET vformdownload = '" & Now & "' WHERE [index] = " & Request("ReqID"), g_strCONN, 1, 3
' Set rsTS = Nothing
' Response.write "ID: " & Request("ReqID")
' Call SaveHist(Request("ReqID"), "[interpreter]dl.asp") 

'downloadfile
'Set dload = Server.CreateObject("SCUpload.Upload")
'tmpfile = attachPDF ' "C:\work\LSS-LBIS\web-DMZ\Images\icon_download.gif" 'attachPDF
'dload.Download tmpFile
'Set dload = Nothing
'download = 1

Set objStream = Server.CreateObject("ADODB.Stream")
objStream.Type = 1 'adTypeBinary
objStream.Open
objStream.LoadFromFile(attachPDF)
Response.ContentType = "application/pdf"
Response.Addheader "Content-Disposition", "attachment; filename=""" & fname  & """"
'Response.Addheader "Content-Length", objStream.Size
Response.BinaryWrite objStream.Read
''objStream.SaveToFile tmpFile, 2
objStream.Close
Set objStream = Nothing
%>