<%language=vbscript%>
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
'GET REQUEST
Set rsConfirm = Server.CreateObject("ADODB.RecordSet")
sqlConfirm = "SELECT * FROM Request_T WHERE [index] = " & Request("ReqID")
rsConfirm.Open sqlConfirm, g_strCONN, 3, 1
If Not rsConfirm.EOF Then
	'TS = rsConfirm("timestamp")
	'RP = rsConfirm("reqID") 
	'tmpClient = ""
	'If rsConfirm("client") = True Then tmpClient = " (LSS Client)"
	'tmpName = rsConfirm("clname") & ", " & rsConfirm("cfname") & tmpClient
	If rsConfirm("CliAdd") = True Then tmpDeptaddr = rsConfirm("caddress") & ", " & rsConfirm("cCity") & ", " &  rsConfirm("cstate") & ", " & rsConfirm("czip")
	'tmpFon = rsConfirm("Cphone")
	'tmpAFon = rsConfirm("CAphone")
	'tmpDir = rsConfirm("directions")
	'tmpSC = rsConfirm("spec_cir")
	'tmpDOB = rsConfirm("DOB")
	'tmpLang = rsConfirm("langID")
	'tmpAppDate = rsConfirm("appDate")
	'tmpAppTFrom = rsConfirm("appTimeFrom") 
	'tmpAppTTo = rsConfirm("appTimeTo")
	'tmpAppLoc = rsConfirm("appLoc")
	'tmpInst = rsConfirm("instID")
	tmpDept = rsConfirm("DeptID")
	'tmpInstRate = Z_FormatNumber(rsConfirm("InstRate"), 2)
	'tmpDoc = rsConfirm("docNum")
	'tmpCRN = rsConfirm("CrtRumNum")
	tmpIntr = rsConfirm("IntrID")
	'tmpIntrRate = Z_FormatNumber(rsConfirm("IntrRate"), 2)
	'tmpEmer = ""
	'If rsConfirm("Emergency") = True Then tmpEmer = "(EMERGENCY)" 
	'tmpCom = rsConfirm("Comment")
	'Statko = GetMyStatus(rsConfirm("Status"))
	'tmpBilHrs = rsConfirm("Billable")
	'tmpActTFrom = Z_FormatTime(rsConfirm("astarttime")) 
	'tmpActTTo = Z_FormatTime(rsConfirm("aendtime"))
	'tmpBilTInst = rsConfirm("TT_Inst")
	'tmpBilTIntr = rsConfirm("TT_Intr")
	'tmpBilMInst = rsConfirm("M_Inst")
	'tmpBilMIntr = rsConfirm("M_Intr")
	'timestamp on sent/print
	'tmpSentReq = "Request email has not been sent to Requesting Person."
	'If rsConfirm("SentReq") <> "" Then tmpSentReq = "Request email was last sent to Requesting Person on <b>" & rsConfirm("SentReq") & "</b>."
	'tmpSentIntr = "Request email has not been sent to Interpreter."
	'If rsConfirm("SentIntr") <> "" Then tmpSentIntr = "Request email was last sent to Interpreter on <b>" & rsConfirm("SentIntr") & "</b>."
	'tmpPrint = "Request has not been printed."
	'If rsConfirm("Print") <> "" Then tmpPrint = "Request was last printed on<b> " & rsConfirm("Print") & "</b>."
	'tmpHPID = Z_CZero(rsConfirm("HPID"))
End If
rsConfirm.Close
Set rsConfirm = Nothing
'GET DEPARTMENT
Set rsDept = Server.CreateObject("ADODB.RecordSet")
sqlDept = "SELECT * FROM dept_T WHERE [index] = " & tmpDept
rsDept.Open sqlDept, g_strCONN, 3, 1
If Not rsDept.EOF Then
	'tmpClass = rsDept("Class")
	'tmpClassName = GetClass(rsDept("Class"))
	'If rsDept("dept") <> "" Then  tmpIname = tmpIname & " - " & rsDept("dept")
	If tmpDeptaddr = "" Then tmpDeptaddr = rsDept("address") & ", " & rsDept("City") & ", " &  rsDept("state") & ", " & rsDept("zip")
	'tmpBaddr = rsDept("Baddress") & ", " & rsDept("BCity") & ", " &  rsDept("Bstate") & ", " & rsDept("Bzip")
	'tmpBContact = rsDept("Blname")
End If
rsDept.Close
Set rsDept = Nothing 
'GET INTERPRETER INFO
Set rsIntr = Server.CreateObject("ADODB.RecordSet")
sqlIntr = "SELECT * FROM interpreter_T WHERE [index] = " & tmpIntr
rsIntr.Open sqlIntr, g_strCONN, 3, 1
If Not rsIntr.EOF Then
	'tmpInName = rsIntr("last name") & ", " & rsIntr("first name")
	'tmpInEmail = rsIntr("E-mail")
	'tmpInFon = rsIntr("phone1")
	'If rsIntr("phone2") <> "" Then tmpInFon = tmpInFon & " / " & rsIntr("phone2")
	'tmpInFax = rsIntr("fax")
	tmpInaddr = rsIntr("address1") & ", " & rsIntr("City") & ", " &  rsIntr("state") & ", " & rsIntr("zip code")	
	'tmpInHouse = ""
	'If rsIntr("InHouse") = True Then tmpInHouse = " (In-House Interpreter)"
End If
rsIntr.Close
Set rsIntr = Nothing
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
    "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"  xmlns:v="urn:schemas-microsoft-com:vml">
  <head>
    <meta http-equiv="content-type" content="text/html; charset=utf-8"/>
    <title>Google Maps for LanguageBank</title>
    <script src=" http://maps.google.com/?file=api&amp;v=2.x&amp;key=<%=googlekey%>"
      type="text/javascript"></script>
    <style type="text/css">
      body {
        font-family: Verdana, Arial, sans serif;
        font-size: 11px;
        margin: 2px;
      }
      table.directions th {
	background-color:#EEEEEE;
      }
	  
      img {
        color: #000000;
      }
    </style>
    <link href='style.css' type='text/css' rel='stylesheet'>
    <script type="text/javascript">
 
    var map;
    var gdir;
    var geocoder = null;
    var addressMarker;
    var duree;

    function initialize() {
      if (GBrowserIsCompatible()) {      
        map = new GMap2(document.getElementById("map_canvas"));
        //gdir = new GDirections(map, document.getElementById("directions"));
        gdir = new GDirections(map, document.getElementById("directions"));
        GEvent.addListener(gdir, "load", onGDirectionsLoad);
        GEvent.addListener(gdir, "error", handleErrors);
				// geocoder = new GClientGeocoder();
				 
        setDirections("<%=tmpInaddr%>", "<%=tmpDeptaddr%>", "en_US");
      }
    }
    
    function setDirections(fromAddress, toAddress, locale) {
    	
      gdir.load("from: " + fromAddress + " to: " + toAddress,
                { "locale": locale });
      }
		
    function handleErrors(){
	   if (gdir.getStatus().code == G_GEO_UNKNOWN_ADDRESS)
	     alert("No corresponding geographic location could be found for one of the specified addresses. This may be due to the fact that the address is relatively new, or it may be incorrect.\nError code: " + gdir.getStatus().code);
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
			 //duree = gdir.getDuration();
				//alert('Dur returned: ' + duree);
			//	document.getElementById("TravT").innerHTML = duree.html; 
      // e.g.
      // document.getElementById("getStatus").innerHTML = gdir.getStatus().code;
	  // and yada yada yada...
	}
    </script>

  </head>
  <body onload="" onunload="GUnload()" bgcolor='#FBF5DB' style="width:100%;height:100%;filter: progid:DXImageTransform.Microsoft.gradient(startColorstr=#FFFFFFF, endColorstr=#FBF5DB);">
  
  <form action="#" onsubmit='setDirections(this.from.value, this.to.value, "en_US"); return false' name='frmtest'>

  <table>
  	<tr>
	<td class='header' colspan='6'>
						<nobr>Directions --&gt&gt
					</td>
		</tr>
		<tr><Td>&nbsp;</td></tr>
   <tr><th align="right">From:&nbsp;</th>

   <td><input type="text" size="35" id="fromAddress" name="from"
     value="<%=tmpInaddr%>"/></td>
   <th align="right">&nbsp;&nbsp;To:&nbsp;</th>
   <td align="right"><input type="text" size="35" id="toAddress" name="to"
     value="<%=tmpDeptaddr%>" /></td></tr>
		
   </table>

    
  </form>

    <br/>
    <table class="directions">
    <tr><th>Formatted Directions</th><th>Map</th></tr>

    <tr>
    <td valign="top"><div id="directions" style="width: 275px"></div></td>
    <td valign="top"><div id="map_canvas" style="width: 310px; height: 400px"></div></td>

    </tr>
    
    </table> 
  </body>
</html>
<script language='JavaScript'><!--
	initialize();
	
-->
</script>