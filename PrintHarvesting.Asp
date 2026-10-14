<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"
%>
<html>
<head>
<title>RW Plant</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js"></script>
<script language="JavaScript" src="date-picker.js"></script> 
</head>
<body>
 <%
    Dim txtNoWS,txtDateWS,txtRemarks,q,nCountofField,no,i
    Dim Remarks
    Dim rsWS,intRecordsPerPage,cColor,txtVehicle
	txtVehicle =cmdFindTitle("Vehicle","fcCode",request("VehicleCode"))
    %>
	 <table width="700">
	  <tr> 
		<td rowspan="6" width="100">
	<% if len(trim(PROFIL("REPORT_IMAGE"))) > 0 then %>
		<img src="IMAGES/<%=PROFIL("REPORT_IMAGE")%>" width="210" height="80">
	<% end if %>
	</td>
	  </tr>
		<tr><td width="100">Company</td><td  width="300"><b><FONT color="#000099"><%=Session("COMPANY")%></FONT></td></tr>
		<tr><td width="100">Business Unit</td><td  width="300"><b><FONT color="#000099"><%=Session("BA")%> / <%=Session("FCBA")%></FONT></td></tr>
		<tr><td width="100">Manager</td><td  width="300"><b><FONT color="#000099"><%=Session("FCMANAGER")%></FONT></td></tr>
		<tr><td width="100">Report</td><td  width="300"><b><FONT color="#000099">HARVESTING</FONT></td></tr>
		<tr><td width="100">For Date</td><td  width="300"><b><FONT color="#000099"><%=FormatDateTime(REQUEST("PeriodDate"),1)%></FONT></td></tr>
	 </table>
	<BR><BR>


   <%
   Dim rsVehicle
   set rsVehicle = server.CreateObject("Adodb.recordset")
   q="select CHITNO, FIELDCODE, CROPCODE, PRODUCTCODE, to_char(RECEPTIONDATE,'mm/dd/yyyy') as RECEPTIONDATE, VEHICLE, OWN, NVL(SUM(QUANTITY),0) AS QUANTITY, " & _
     "NVL(SUM(ESTATE_WEIGHT),0) AS ESTATE_WEIGHT, NVL(SUM(MILL_WEIGHT),0) AS MILL_WEIGHT from Harvesting where " & _
     " to_char(RECEPTIONDATE,'mm/yyyy')=to_char(to_date('" & request("PeriodDate") & "','" & "MM/DD/YYYY" & "'),'mm/yyyy') and fcBA='" & session("fcBA") & "'"

	if  len(trim(request("PeriodDate"))) > 0 then
		q = q & " and to_char(RECEPTIONDATE,'mm/yyyy')=to_char(to_date('" & request("PeriodDate") & "','" & "MM/DD/YYYY" & "'),'mm/yyyy')"
	end if
	if  len(trim(request("Txt_Fieldcode"))) > 0 then
		q = q & " and FIELDCODE='" & request("Txt_Fieldcode") & "'"
	end if
	q = q & " group by CHITNO, FIELDCODE, CROPCODE, PRODUCTCODE, to_char(RECEPTIONDATE,'mm/dd/yyyy'), VEHICLE, OWN"
  
   'response.Write (q)
   'response.Flush()	
   Set rsVehicle = Server.CreateObject("ADODB.Recordset")
   rsVehicle.CursorType = 3
   rsVehicle.CursorLocation = 3
   rsVehicle.Open q, connstr%>
   <br><br>
	<table Cellpadding="2" bgcolor="#F7F7F7" cellspacing="0">
   <tr>   
	    <td style="<%=cmdstyleLeft%>;border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"><font face="Verdana" size="1">NO</td>	   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">CHITNO</td></td>
	   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">FIELD</td>
	   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">CROP</td>
	   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">PRODUCT</td>
	   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">RECEPTIONDATE</td>
	   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">VEHICLE</td>
	   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">OWN</td>
	   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">QUANTITY</td>
	   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">ESTATEWEIGHT</td>
	   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">MILLWEIGHT</td>
   </tr>
   <%
   Dim nMillweight,nEstate,nQuantity
   if Not rsVehicle.eof then
      nCountofField=rsVehicle.fields.count-1
	  no=1
		nMillweight=0
		nEstate=0
		nQuantity=0
	  While Not rsVehicle.EOF
		  nMillweight = nMillweight+cdbl(rsVehicle("MILL_WEIGHT"))
		  nEstate = nEstate+cdbl(rsVehicle("ESTATE_WEIGHT"))
		  nQuantity = nQuantity+cdbl(rsVehicle("QUANTITY"))
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		<td align="left" style="<%=cmdstyleLeft%>"><font face="Verdana" size="1"><%=NO%>&nbsp;</td>
          <td align="left"  title="CHITNO" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsVehicle("CHITNO")%>&nbsp;</td>
	      <td align="left"  title="FIELDCODE" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsVehicle("FIELDCODE")%>&nbsp;</td>
	      <td align="left"  title="CROPCODE" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsVehicle("CROPCODE")%>&nbsp;</td>
	      <td align="left"  title="PRODUCTCODE" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsVehicle("PRODUCTCODE")%>&nbsp;</td>
	      <td align="left"  title="RECEPTIONDATE" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsVehicle("RECEPTIONDATE")%>&nbsp;</td>
	      <td align="left"  title="VEHICLE" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsVehicle("VEHICLE")%>&nbsp;</td>
	      <td align="left"  title="VEHICLE" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsVehicle("OWN")%>&nbsp;</td>		  
	      <td align="right" title="WEIGHT" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsVehicle("QUANTITY"),2)%>&nbsp;</td>
	      <td align="right" title="QUANTITY" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsVehicle("ESTATE_WEIGHT"),2)%>&nbsp;</td>
	      <td align="right"  title="COLLECTIONDATE" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(rsVehicle("MILL_WEIGHT"),2)%>&nbsp;</td>
		  </tr>
		  <%
		  rsVehicle.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend%>
	  <tr>
	  <td colspan="8" align="center" style="<%=cmdstyleLeft%>"><font face="Verdana" size="1">T o t a l</td>
	  <td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(nQuantity,2)%>&nbsp;</td>
	  <td align="right"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(nEstate,2)%>&nbsp;</td>
	  <td align="right"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(nMillweight,2)%>&nbsp;</td>
	  </tr><%
   else
      response.Write("Data is empty")	  
   end if 
   rsVehicle.Close
   set rsVehicle = Nothing%>
   </TABLE>
<!--#include file="Bottom_Report.asp"-->