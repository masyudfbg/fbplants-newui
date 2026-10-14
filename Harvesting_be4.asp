<%Option Explicit%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
<head>
<title>ePlaS - Harvesting Receptions</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
<script type="text/javascript" src="dtree.js"></script>
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
<script language="JavaScript">
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window);
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);
</script>
 <script>
$(function() {
$( ".datepicker" ).datepicker();
});
</script>
<!--<script language="JavaScript" src="/js/date-picker.js"></script> -->
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>
</head>
<body>
<!--#include file="top.jsp"-->
	<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
			<!--#include file="menu.asp"-->
		</div>
   <div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
 <div class="judul">
	   Module : <%=request("txtModule")%> / <%=request("txtTable")%>
	</div>
 <div class="box_both" style="overflow:scroll;">
 <!-------------awal!------------->
 <%
 Dim FCCODE
 Dim FCENTRY, FCEDIT, FCIP, FCBA
 Dim CHITNO, FIELDCODE, CROPCODE, PRODUCTCODE, RECEPTIONDATE, VEHICLE, OWN, TRANSPORTTYPE, CONTRACTNUMBER, QUANTITY, ESTATE_WEIGHT, MILL_WEIGHT
 Dim cmdView

RECEPTIONDATE =request("header_date")
CHITNO=request("CHITNO")
FIELDCODE=request("FIELDCODE") 
CROPCODE=request("CROPCODE")
PRODUCTCODE=request("ProductCODE")
RECEPTIONDATE=request("RECEPTIONDATE")
VEHICLE=request("VEHICLE")
OWN=request("OWN")
TRANSPORTTYPE = Request("TRANSPORTTYPE")
CONTRACTNUMBER = Request("CONTRACTNUMBER")
QUANTITY=request("QUANTITY")
ESTATE_WEIGHT=request("ESTATE_WEIGHT")
MILL_WEIGHT=request("MILL_WEIGHT")
 
if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
	if Request("txtJenis")= "Delete" then
		q	=	" Delete from Harvesting where" & _
				" CHITNO = '" & CHITNO & "'" & _
				" And FIELDCODE = '" & FIELDCODE & "'" & _
				" And RECEPTIONDATE = to_date('" & RECEPTIONDATE & "','mm/dd/yyyy')" & _
				" and fcBA = '" & Session("fcBA") & "'"

		'Response.Write q
		connStr.execute(q)
		cmdView	= "True"
	elseif Request("txtJenis")="New" then
	If CheckIsTM(FIELDCODE, Session("fcBA")) > 0 Then
		q	=	" Insert into Harvesting (" & _
				" CHITNO" & _
				" , FIELDCODE" & _
				" , CROPCODE" & _
				" , PRODUCTCODE" & _
				" , RECEPTIONDATE" & _
				" , VEHICLE" & _
				" , OWN" & _
				" , QUANTITY" & _
				" , ESTATE_WEIGHT" & _
				" , MILL_WEIGHT" & _
				" , FCENTRY" & _
				" , FCEDIT" & _
				" , FCIP" & _
				" , FCBA" & _
				" , LASTUPDATE" & _
				" , LASTTIME" & _
				" ) values (" & _
				" '" & CHITNO & "'" & _
				" , '" & FIELDCODE & "'" & _
				" , '" & CROPCODE & "'" & _
				" , '" & PRODUCTCODE & "'" & _
				" , to_date('" & RECEPTIONDATE & "','mm/dd/yyyy')" & _
				" , '" & VEHICLE & "'" & _
				" , '" & OWN & "'" & _
				" , '" & QUANTITY & "'" & _
				" , '" & ESTATE_WEIGHT & "'" & _
				" , '" & MILL_WEIGHT & "'" & _
				" , '" & Session("CodeName") & "'" & _
				" , '" & Session("CodeName") & "'" & _
				" , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
				" , '" & Session("fcBA") & "'" & _
				" , sysdate " & _
				" , To_Char(sysdate, 'HH24:MI')" & _
				" )"
				
		'Response.Write q
		connStr.execute(q)
		Else
			err.description = "Blok " & FIELDCODE & " berstatus TBM"
			err.number = 10001
		End If
		
		cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		q	=	" Update Harvesting set" & _
				" CROPCODE = '" & CROPCODE & "'" & _
				" , PRODUCTCODE = '" & PRODUCTCODE & "'" & _
				" , VEHICLE = '" & VEHICLE & "'" & _
				" , OWN = '" & OWN & "'" & _
				" , QUANTITY = '" & QUANTITY & "'" & _
				" , ESTATE_WEIGHT = '" & ESTATE_WEIGHT & "'" & _
				" , MILL_WEIGHT = '" & MILL_WEIGHT & "'" & _
				" , LASTUPDATE = sysdate " & _
		        " , LASTTIME = To_Char(sysdate,'HH24:MI') " & _
				" where" & _
				" CHITNO = '" & CHITNO & "'" & _
				" And FIELDCODE = '" & FIELDCODE & "'" & _
				" And RECEPTIONDATE = to_date('" & RECEPTIONDATE & "','mm/dd/yyyy')" & _
				" and fcBA = '" & Session("fcBA") & "'"
		
		'Response.Write q
		connStr.execute(q)
		connStr.execute("commit")
		cmdView	= "True"
    else
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
else
    cmdView	= request("cmdView")
End if
Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis
if RECEPTIONDATE <> "" then
   RECEPTIONDATE   =Request("RECEPTIONDATE")
else
   RECEPTIONDATE   =textDate(date)
end if
txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
txtJenis = request("txtJenis")

If Err.Number <> 0 then
   Response.Write "<font color=red>" & "Error Number : " & Err.Number &  " Desc. " & Err.Description & "</font>"
End if
%>
 <!-------------ISI Payment Form!------------->
 <form name="frmEntry" method="post" action="Harvesting.asp">
	<table class="data" width="70%" cellpadding="0" cellspacing="0">
		<tr>
			<td colspan="2">
				<div align="left"><font size="2" face="Verdana"><strong>MAINTENANCE HARVESTING</strong></font></div>
			</td>
		</tr>

		<tr>
			<td align="left" width="30%">Date</td>
			<td align="left">
				<input class="datepicker" type="text" maxlength="10" name="RECEPTIONDATE" size="10" value="<%=RECEPTIONDATE%>">
				
			</td>
		</tr>
			
		<tr>
			<td colspan="2"  align="left">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit" name="Add" value="Select">&nbsp;&nbsp;<a href="PrintHarvesting.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&RECEPTIONDATE=<%=RECEPTIONDATE%>" title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>
			</td>
		</tr>
    </table>  
 </form>
<%if cmdView = "True" And len(trim(RECEPTIONDATE)) > 0 Then
   Dim rsVH,intRecordsPerPage,cColor
   set rsVH = server.CreateObject("Adodb.recordset")
   'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
   q	=	" select CHITNO, FIELDCODE, CROPCODE, ProductCODE, to_char(RECEPTIONDATE,'mm/dd/yyyy') as RECEPTIONDATE, VEHICLE, OWN, QUANTITY, " & _
			" ESTATE_WEIGHT, MILL_WEIGHT, FCENTRY, FCEDIT, FCIP, FCBA from Harvesting where fcBA='" & session("fcBA") & "'" & _
			" and RECEPTIONDATE = to_date('" & RECEPTIONDATE & "', 'MM/DD/YYYY')" & _
			" ORDER BY CHITNO, FIELDCODE"
	
	Set rsVH = Server.CreateObject("ADODB.Recordset")
	'response.Write(q)
	'response.Flush()
	rsVH.CursorType = 3
	rsVH.CursorLocation = 3
	rsVH.Open q, connstr
   
%>
   <form action="Harvesting.Asp" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0" width="70%">
		
		<tr>
			<td align="left" width="30%">Chit No</td>
			<td align="left">
				<input type="text" size="15" name="CHITNO" value="<%=CHITNO%>" <%IF txtJenis="Edit" Then%> readonly="True" style="background:#999999"<%End if%>><font color="#FF0000"> *</font>
			</td>
		</tr>
		
		<tr>
			<td align="left">Field Code</td>
			<td align="left">
				<select name="FIELDCODE" <%IF Request("txtJenis")="Edit" then%> disabled = "disabled" <%End If%>>
					<%cmdCallCombo "Field",FIELDCODE,"",""%>
				</select><font color="#FF0000"> *</font>
				<%IF Request("txtJenis")="Edit" then%> 
				<input type="hidden" name="FIELDCODE" value="<%=FIELDCODE%>" />
				<%End If%>
			</td>
		</tr>
		
		<tr>
			<td align="left">Crop Code</td>
			<td align="left">
				<select name="CROPCODE">
					<option value=""></option> 
				<%
					Dim rsCrop, strCrop
					strCrop =	" SELECT P.FCCODE,P.FCNAME FROM  PARAMETER P  WHERE P.PARCODE= 'F0005' AND FCBA = '" &  Session("fcBA") & "'" & _
								" ORDER BY P.FCNAME DESC"
					
					Set rsCrop = Server.CreateObject("ADODB.Recordset")
					rsCrop.Open strCrop, connstr
					
					If Not rsCrop.EOF then
						Do While Not rsCrop.EOF
				%>
					<option value="<%=rsCrop("FCCODE")%>" <%If CROPCODE = rsCrop("FCCODE") Then%> selected <%End If%>><%=rsCrop("FCNAME")%></option> 
				<%			rsCrop.movenext
						Loop
					End if
					rsCrop.Close
					set rsCrop = Nothing
				%>
				</select><font color="#FF0000"> *</font>
			</td>
		</tr>
		
		<tr>
			<td align="left"><font face="Verdana" size="1">Product Code</font></td>
			<td align="left">
				<select name="ProductCODE">
					<option value=""></option> 
				<%
					Dim rsProductCode, strProductCode
					strProductCode = " SELECT PRODUCTCODE, PRODUCTDESCRIPTION FROM PRODUCT WHERE FCBA = '" &  Session("fcBA") & "' ORDER BY PRODUCTCODE"
					
					Set rsProductCode = Server.CreateObject("ADODB.Recordset")
					rsProductCode.Open strProductCode, connstr
					
					If Not rsProductCode.EOF then
						Do While Not rsProductCode.EOF
				%>
					<option value="<%=rsProductCode("PRODUCTCODE")%>" <%If PRODUCTCODE = rsProductCode("PRODUCTCODE") Then%> selected <%End If%>><%=rsProductCode("PRODUCTCODE")%> : <%=rsProductCode("PRODUCTDESCRIPTION")%></option> 
				<%			rsProductCode.movenext
						Loop
					End if
					rsProductCode.Close
					set rsProductCode = Nothing
				%>
				</select><font color="#FF0000"> *</font>
			</td>
		</tr>

		<tr>
			<td align="left"><font face="Verdana" size="1">Vehicle Status</font></td>
			<td align="left">
				<select name="OWN">
				<%
					Dim rsVehicleStatus, strVehicleStatus
					strVehicleStatus = " SELECT FCCODE, FCNAME FROM HARVESTINGVEHICLESTATUS ORDER BY FCCODE"
					
					Set rsVehicleStatus = Server.CreateObject("ADODB.Recordset")
					rsVehicleStatus.Open strVehicleStatus, connstr
					
					If Not rsVehicleStatus.EOF then
						Do While Not rsVehicleStatus.EOF
				%>
					<option value = "<%=rsVehicleStatus("FCCODE")%>" <%If OWN = rsVehicleStatus("FCCODE") Then%> selected <%End If%>>
						<%=rsVehicleStatus("FCNAME")%>
					</option> 
				<%			rsVehicleStatus.movenext
						Loop
					End if
					rsVehicleStatus.Close
					set rsVehicleStatus = Nothing
				%>
				</select><font color="#FF0000"> *</font>
			</td>
		</tr>
		
		<tr>
            <td align="left">Vehicle</td>
			<td align="left">
				<input type="text" size="15" maxlength="50" name="Vehicle" value="<%=Vehicle%>"><font color="#FF0000"> *</font>
			</td>
		</tr>
		
		<tr>
			<td align="left"><font face="Verdana" size="1">Transport Type</font></td>
			<td align="left">
				<select name="TRANSPORTTYPE">
				<%
					Dim rsTransportType, strTransportType
					strTransportType = " SELECT FCCODE, FCNAME FROM HARVESTINGTRANSPORTTYPE ORDER BY FCCODE"
					
					Set rsTransportType = Server.CreateObject("ADODB.Recordset")
					rsTransportType.Open strTransportType, connstr
					
					If Not rsTransportType.EOF then
						Do While Not rsTransportType.EOF
				%>
					<option value = "<%=rsTransportType("FCCODE")%>" <%If TRANSPORTTYPE = rsTransportType("FCCODE") Then%> selected <%End If%>>
						<%=rsTransportType("FCNAME")%>
					</option> 
				<%			rsTransportType.movenext
						Loop
					End if
					rsTransportType.Close
					set rsTransportType = Nothing
				%>
				</select><font color="#FF0000"> *</font>
			</td>
		</tr>
		
		<tr>
            <td align="left">Contract Number</td>
			<td align="left">
				<input type="text" size="30" maxlength="30" name="CONTRACTNUMBER" value="<%=CONTRACTNUMBER%>">
			</td>
		</tr>

		<tr><td align="left">Bunch</td><td align="left"><input type="text" size="15" name="QUANTITY" value="<%=QUANTITY%>"></td></tr>
		<tr><td align="left">Estate Weight (Kg)</td><td align="left"><input type="text" size="15" name="ESTATE_WEIGHT" value="<%=ESTATE_WEIGHT%>"></td></tr>
		<tr><td align="left">Mill Weight (Kg)</td><td align="left"><input type="text" size="15" name="MILL_WEIGHT" value="<%=MILL_WEIGHT%>"></td></tr>
	    <TR><td colspan="2" align="left">
		  <input type="hidden" name="RECEPTIONDATE" value="<%=RECEPTIONDATE%>">
		  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		  <%IF Request("txtJenis")="Edit" then%>
		       <input type="hidden" name="txtJenis" value="Edit">
		       <input Type="Submit"  name="Add" value="Update">
		  <%Else%>
		       <input type="hidden" name="txtJenis" value="New">
		       <input Type="Submit"  name="Add" value="Add">
		  <%End if%>
		  <input Type="Reset"  name="Reset" value="Reset">
		  <input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
	  </tr>
   </table>
 </form>
  <!--#include file="validation_show.Asp"--> 
   <table class="data">
	<BR>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">CHIT NO</TD>	
	<TD align="left">FILED CODE</TD>	
	<TD align="left">CROP CODE</TD>	
	<TD align="left">PRODUCT CODE</TD>	
	<TD align="left">RECEPTION DATE</TD>	
	<TD align="left">VEHICLE</TD>	
	<TD align="left">OWN</TD>	
	<TD align="left">BUNCH</TD>	
	<TD align="left">ESTATE WEIGHT</TD>	
	<TD align="left">MILL WEIGHT</TD>	
	</tr>
   <%
   if Not rsVH.eof then
      nCountofField=rsVH.fields.count-1
	  no=1
	  While Not rsVH.EOF
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td align="left" valign="top"><%=NO%></td>
			<td valign="middle" width="30">
			<a href="Harvesting.Asp?CHITNO=<%=rsVH("CHITNO")%>&FIELDCODE=<%=rsVH("FIELDCODE")%>&CROPCODE=<%=rsVH("CROPCODE")%>&PRODUCTCODE=<%=rsVH("PRODUCTCODE")%>&RECEPTIONDATE=<%=rsVH("RECEPTIONDATE")%>&VEHICLE=<%=rsVH("VEHICLE")%>&OWN=<%=rsVH("OWN")%>&QUANTITY=<%=rsVH("QUANTITY")%>&ESTATE_WEIGHT=<%=rsVH("ESTATE_WEIGHT")%>&MILL_WEIGHT=<%=rsVH("MILL_WEIGHT")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsVH(1)%>" border="0"></a>
			</td>
			<td valign="middle" width="30">
			<a href="Harvesting.Asp?CHITNO=<%=rsVH("CHITNO")%>&FIELDCODE=<%=rsVH("FIELDCODE")%>&RECEPTIONDATE=<%=rsVH("RECEPTIONDATE")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsVH(1)%>" border="0" onClick="return confirm('Are you sure to Delete');"></a></td>
		    <%for i = 0 to nCountofField-4%>
			    <td align="left" title="<%=rsVH(i).Name%>" valign="top"><%=rsVH(i)%></td>
		    <%next%>
		  </tr>
		  <%
		  rsVH.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   end if 
   rsVH.Close
   set rsVH = Nothing%>
   </TABLE>
  </div>

<%End if%>  
 
  <!-------------akhir!------------->
</div></div>
  <!--#include file="bottom.jsp"-->

</body>
</html>
