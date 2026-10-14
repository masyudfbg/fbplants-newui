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
		<tr><td width="100">Report</td><td  width="300"><b><FONT color="#000099">DELIVERY CONTRACT</FONT></td></tr>
		<tr><td width="100">For Date</td><td  width="300"><b><FONT color="#000099"><%=FormatDateTime(REQUEST("PeriodDate"),1)%></FONT></td></tr>
	 </table>
	<BR><BR>


   <%
   Dim rsVehicle
   set rsVehicle = server.CreateObject("Adodb.recordset")
   q="select CONTRACTCODE, CUST_CUSTOMERCODE, STARTDATE, ENDDATE, ACTUALSTART, ACTUALEND, " & _
     "PRODUCT_PRODUCTCODE, TOTALWEIGHT, STATUS, REFERENCE, FCENTRY, FCEDIT, FCIP, FCBA from DELIVERYCONTRACT where " & _
     " fcBA='" & session("fcBA") & "'"
   Set rsGAD = Server.CreateObject("ADODB.Recordset")
	if  len(trim(request("txt_CONTRACTCODE"))) > 0 then
		q = q & " and CONTRACTCODE='" & request("txt_CONTRACTCODE") & "'"
	end if
	q = q


  
   'response.Write (q)
   'response.Flush()	
   Set rsVehicle = Server.CreateObject("ADODB.Recordset")
   rsVehicle.CursorType = 3
   rsVehicle.CursorLocation = 3
   rsVehicle.Open q, connstr%>
   <br><br>
	<table Cellpadding="2" bgcolor="#F7F7F7" cellspacing="0">
   <tr>   
	    <td style="<%=cmdstyleLeft%>;border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"><font face="Verdana" size="1">NO</td>	 
	   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">CODE</td>
	   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">CUSTOMER</td>
	   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">START DATE</td>
	   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">END DATE</td>
	   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">ACTUAL START</td>
	   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">ACTUAL END</td>
	   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">PRODUCT CODE</td>
	   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">TOTAL WEIGHT</td>
	   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">STATUS</td>
	   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">REFERENCES</td>
   </tr>
   <%
   if Not rsVehicle.eof then
      nCountofField=rsVehicle.fields.count-1
	  no=1
		nMillweight=0
		nEstate=0
		nQuantity=0
	  While Not rsVehicle.EOF
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		<td align="left" style="<%=cmdstyleLeft%>"><font face="Verdana" size="1"><%=NO%>&nbsp;</td>
          <td align="left"  title="PRODUCTCODE" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsVehicle("CONTRACTCODE")%>&nbsp;</td>
	      <td align="left"  title="PRODUCTDESCRIPTION" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsVehicle("CUST_CUSTOMERCODE")%>&nbsp;</td>
	      <td align="left"  title="CROPCODE" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsVehicle("STARTDATE")%>&nbsp;</td>
	      <td align="left"  title="UNITOFMEASURE" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsVehicle("ENDDATE")%>&nbsp;</td>
	      <td align="right"  title="FORMULA" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsVehicle("ACTUALSTART")%>&nbsp;</td>
	      <td align="right"  title="FORMULA" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsVehicle("ACTUALEND")%>&nbsp;</td>
	      <td align="right"  title="FORMULA" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsVehicle("PRODUCT_PRODUCTCODE")%>&nbsp;</td>
	      <td align="right"  title="FORMULA" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsVehicle("TOTALWEIGHT")%>&nbsp;</td>
	      <td align="right"  title="FORMULA" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsVehicle("STATUS")%>&nbsp;</td>
	      <td align="right"  title="FORMULA" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsVehicle("REFERENCE")%>&nbsp;</td>
		  
		  </tr>
		  <%
		  rsVehicle.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend%>
<%
   else
      response.Write("Data is empty")	  
   end if 
   rsVehicle.Close
   set rsVehicle = Nothing%>
   </TABLE>
<!--#include file="Bottom_Report.asp"-->