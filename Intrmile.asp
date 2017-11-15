<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
If Left(Request.ServerVariables("REMOTE_ADDR"), 10) = "192.168.1." Or _
	Request.ServerVariables("REMOTE_ADDR") = "127.0.0.1" Or _
	Left(Request.ServerVariables("REMOTE_ADDR"), 8) = "10.10.1." Then
	googlekey = "ABQIAAAAd5OxJhCEqwRNwElUvBNmZxR9PeFMte5gUE1Dq7em5JwYVo_dVhScQdsXHPRROmqe71rlFsfMGLuovg"
Else
	googlekey = "ABQIAAAAd5OxJhCEqwRNwElUvBNmZxSZl_t-SL-f-oE8q1L92qagyvYqqhSeaBa4qBIqCn9H6Ik6hSNkS-Lp6w"
End If
	Function GetDeptZip(xxx)
		Set rsDept = Server.CreateObject("ADODB.RecordSet")
		sqlDept = "SELECT * FROM dept_T WHERE [index] = " & xxx
		rsDept.Open sqlDept, g_strCONN, 3, 1
		If Not rsDept.EOF Then
			GetDeptZip = rsDept("zip")
		End If
		rsDept.Close
		Set rsDept = Nothing 
	End Function
	'get mileage cap for interpreters
	set rsmile = server.createobject("adodb.recordset")
	sqlmile = "select * from travel_t"
	rsmile.open sqlmile, g_strconn, 3, 1
	if not rsmile.eof then
		tmpmilecap = Z_czero(rsmile("milediff"))
	end if
	rsmile.close
	set rsmile = nothing
	'GET ADDRESS AND ZIP of Intrpreter
	Set rsIntr = Server.CreateObject("ADODB.REcordSet")
	sqlIntr = "SELECT * FROM Interpreter_T WHERE [index] = " & Request("selIntr")
	rsIntr.Open sqlIntr, g_strCONN, 1, 3
	If Not rsIntr.EOF Then
		tmpIntrAdd = rsIntr("address1") & ", " & rsIntr("City") & ", " &  rsIntr("state") & ", " & rsIntr("Zip Code")
		tmpIntrZip = rsIntr("Zip Code")
	End If
	rsIntr.Close
	Set rsIntr = Nothing
	'GET ADDRESS AND ZIP of DEPARTMENT/CLIENT
	Set rsConfirm = Server.CreateObject("ADODB.RecordSet")
	sqlConfirm = "SELECT * FROM Request_T WHERE [index] = " & Request("RID")
	rsConfirm.Open sqlConfirm, g_strCONN, 3, 1
	If Not rsConfirm.EOF Then
		If rsConfirm("CliAdd") = True Then 
			tmpDeptaddr = rsConfirm("CAddress") &", " & rsConfirm("CCity") & ", " & rsConfirm("CState") & ", " & rsConfirm("CZip")
			tmpZipInst = rsConfirm("czip")
		Else
			tmpDeptaddr = GetDeptAdr(rsConfirm("DeptID"))
			tmpZipInst = GetDeptZip(rsConfirm("DeptID"))
		End If
	End If
	rsConfirm.CLose
	Set rsConfirm = Nothing
%>
<html>
	<head>
		<title>Email Interpreter</title>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script src=" http://maps.google.com/?file=api&amp;v=2.x&amp;key=<%=googlekey%>"
      type="text/javascript"></script>
    <script language='JavaScript'>
    	var map;
	    var gdir;
	    var geocoder = null;
	    var addressMarker;
	    var duree;
	    var dist;
			var dureeHrs;
			var distMile;
			
	    function initialize() {
	      if (GBrowserIsCompatible()) {      
	        map = new GMap2(document.getElementById("map_canvas"));
	        //gdir = new GDirections(map, document.getElementById("directions"));
	        gdir = new GDirections(map, document.getElementById("directions"));
	        GEvent.addListener(gdir, "load", onGDirectionsLoad);
	        GEvent.addListener(gdir, "error", handleErrors);
	        //geocoder = new GClientGeocoder();
					
					
	        setDirections("<%=tmpIntrAdd%>", "<%=tmpDeptaddr%>", "en_US");
	      }
	    }
	    function setDirections(fromAddress, toAddress, locale) {
	    	//geocoder.getLocations(gdir.getGeocode(1), callback);
	      gdir.load("from: " + fromAddress + " to: " + toAddress,
	                { "locale": locale });
	     
	    }
	
	    function handleErrors(){
		   if (gdir.getStatus().code == G_GEO_UNKNOWN_ADDRESS)
		     {
		   		var ans = window.confirm("No corresponding geographic location could be found for one of the specified addresses. This may be due to the fact that the address is relatively new, or it may be incorrect.\nError code: " + gdir.getStatus().code + "\n\nDo you want ZIP CODES to be used instead ?");
		   		if (ans)
		   		{		
		   				//document.frmConfirm.zipcalc.disabled = true;
		   				setDirections("<%=tmpIntrZip%>", "<%=tmpZipInst%>", "en_US");
		   		}
		   		else
	   			{
	   				//document.frmConfirm.zipcalc.disabled = false;
	   			}
		   	}
		   else if (gdir.getStatus().code == G_GEO_SERVER_ERROR)
		     alert("A geocoding or directions request could not be successfully processed, yet the exact reason for the failure is not known.\n Error code: " + gdir.getStatus().code);
		   
		   else if (gdir.getStatus().code == G_GEO_MISSING_QUERY)
		     alert("The HTTP q parameter was either missing or had no value. For geocoder requests, this means that an empty address was specified as input. For directions requests, this means that no query was specified in the input.\n Error code: " + gdir.getStatus().code);
	
		//   else if (gdir.getStatus().code == G_UNAVAILABLE_ADDRESS)  <--- Doc bug... this is either not defined, or Doc is wrong
		//     alert("The geocode for the given address or the route for the given directions query cannot be returned due to legal or contractual reasons.\n Error code: " + gdir.getStatus().code);
		     
		   else if (gdir.getStatus().code == G_GEO_BAD_KEY)
		     alert("The given key is either invalid or does not match the domain for which it was given. \n Error code: " + gdir.getStatus().code);
	
		   else if (gdir.getStatus().code == G_GEO_BAD_REQUEST)
		     alert("A directions request could not be successfully parsed.\n Error code: " + gdir.getStatus().code);
		    
		   else alert("An unknown error occurred.");
		   
		}
	
		function onGDirectionsLoad(){ 
	      // Use this function to access information about the latest load()
	      // results.
				duree = gdir.getDuration();
				dist = gdir.getDistance();
				dureeHrs = ((duree.seconds) / 60) / 60;
				distMile = dist.meters / 1609.344;
				//document.getElementById("ttM").innerHTML = (Math.round(dureeHrs * 100)/100) + " Hrs. - " + (Math.round(distMile*100)/100) + " Miles"; 
	   		decHrs = dureeHrs;
				decMile = distMile;
	   		tmpRate = decMile / decHrs;
	   		//alert(decHrs + "         " + decMile);
	   		if (decMile > <%=tmpmilecap%>) //interpreter
		  	{
		  		bilMile = (decMile * 2) - (<%=tmpmilecap%> * 2); //billable mileage (2 way)
					bilTravel = bilMile / tmpRate; //billable travel time (2 way)
		   		document.frmMile.txtTravel.value = Math.round(bilTravel * 100)/100;
			  	document.frmMile.txtMile.value = Math.round(bilMile * 100)/100;
		  	}
		  	else
		  	{
		  		document.frmMile.txtTravel.value = 0;
		  		document.frmMile.txtMile.value = 0;
		  	}
		  	//alert(document.frmMile.txtTravel.value + "         " + document.frmMile.txtMile.value);
		 	}
		 	function SubmitMe()
		 	{
		 			//alert(document.frmMile.txtTravel.value + "    " + document.frmMile.txtMile.value);
		 			document.frmMile.action = "emailIntr.asp"
		 			document.frmMile.submit();	
		 	}
    </script>
	</head>
	<body onload='initialize();' onunload="GUnload();">
		<form name='frmMile' method='post'>
			<center>
			<table>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td align='right'>
						Intepreter:
					</td>
					<td>
						<b><%=GetIntr(Request("selIntr"))%></b>
					</td>
				</tr>
				<tr>
					<td align='right'>
						Mileage:
					</td>
					<td>
						<input class='main' size='5' readonly name='txtMile'>&nbsp;miles
					</td>
				</tr>
				<tr>
					<td align='right'>
						Travel Time:
					</td>
					<td>
						<input class='main' size='5' readonly name='txtTravel'>&nbsp;hrs
					</td>
				</tr>	
				<tr>
					<td colspan='2' align='center'>
							<input class='btn' type='button' value='OK' style='width: 100px;' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='SubmitMe();'>
								<input class='btn' type='button' value='Back' style='width: 100px;' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='document.location="emailIntr.asp?ID=<%=Request("RID")%>";'>
					</td>
				</tr>
			</table>
			
			<input type='hidden' name='selIntr' value='<%=Request("selIntr")%>'>
			<input type='hidden' name='adr1'  value='<%=tmpIntrAdd%>'>
			<input type='hidden' name='adr2'  value='<%=tmpDeptaddr%>'>
			<input type='hidden' name='zip1'  value='<%=tmpIntrZip%>'>
			<input type='hidden' name='zip2'  value='<%=tmpZipInst%>'>
			<input type='hidden' name='ID'  value='<%=Request("RID")%>'>
			<div id="directions" style="display: none;"></div>
			<div id="map_canvas" style="display: none;"></div>
		</form>
	</body>
</html>
<!-- #include file="_closeSQL.asp" -->
<script language='JavaScript'>
	//initialize();
	
</script>