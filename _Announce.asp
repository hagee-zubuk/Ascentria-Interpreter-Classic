<%
strAnn = ""
Set rsAnn = Server.CreateObject("ADODB.RecordSet")
sqlAnn = "SELECT Intr FROM Announce_T"
rsAnn.Open SqlAnn, g_strCONN, 3, 1
If Not rsAnn.EOF Then
	strAnn = rsAnn("Intr")
End If
rsAnn.Close
Set rsAnn = Nothing
%>