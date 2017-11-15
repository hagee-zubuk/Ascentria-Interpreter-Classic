<%
'header of web page - includes menu
%>
<table cellSpacing='0' cellPadding='0' width="100%" border='0'>
	<tr>
		<td valign='top' align="left" rowspan="2" width="75%" height="65px" colspan="10">
			<img src='images/LBISLOGO.jpg' border="0">
		</td>
		<td align="center" width="25%" class="tollnum">
		Toll-Free 844.579.0610
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>	
	<tr bgcolor='#f68328'>
		<td class="motto" align="center">
			Understand and Be Understood.
		</td>
		<% If Request.Cookies("LBUSERTYPE") <> 2 Then %>
			
		<% Else %>
			<td align='center' class='head' width='10px'>&nbsp;|&nbsp;</td> 
			<td align='center' width='90px'><a href='availappt.asp' class='link2'><nobr>Open Appointments</a></td>
			<td align='center' class='head' width='10px'>&nbsp;|&nbsp;</td>
			<td align='center' width='90px'><a href='calendarview2.asp' class='link2'>Calendar</a></td>
			<td align='center' class='head' width='10px'>&nbsp;|&nbsp;</td>
			<td align='center' width='90px'><a href='avail2.asp' class='link2'>Availability</a></td>
			<td align='center' class='head' width='10px'>&nbsp;|&nbsp;</td> 
			<td align='center' width='70px'><a href='links.asp' class='link2'>Trainings &amp; Links</a></td>
			<td align='center' class='head' width='10px'>&nbsp;|&nbsp;</td> 
			<% If Request.Cookies("ONCALL") = 1 Then %>
				<td align='center' width='90px'><a href='oncall.asp' class='link2'>On Call Sched</a></td>
				<td align='center' class='head' width='10px'>&nbsp;|&nbsp;</td>   
			<% End If %>
		<% End If %>
		<td align='right'><a href='default.asp?chk=1' class='link2'>Sign Out</a>&nbsp;</td>
	</tr>
</table>