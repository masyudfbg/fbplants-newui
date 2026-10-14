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
		<tr><td width="100">Report</td><td  width="300"><b><FONT color="#000099">SUPPLIER</FONT></td></tr>
		<tr><td width="100">For Date</td><td  width="300"><b><FONT color="#000099"><%=FormatDateTime(REQUEST("PeriodDate"),1)%></FONT></td></tr>
	 </table>
	<BR><BR>


   <%
   Dim rsVehicle
   set rsVehicle = server.CreateObject("Adodb.recordset")
   q=" select CUSTOMERCODE, DESCRIPTION, SUPPLIER, ICNO, BANK, BANKACCOUNTNO, COMPANYNO, FCENTRY, FCEDIT, FCIP, FCBA " & _
	 " from CUSTOMER " & _
     " where UPPER(SUPPLIER) ='Y' AND fcBA='" & session("fcBA") & "'"
	 
   Set rsGAD = Server.CreateObject("ADODB.Recordset")
	if  len(trim(request("txt_CUSTOMERCODE"))) > 0 then
		q = q & " and CUSTOMERCODE='" & request("txt_CUSTOMERCODE") & "'"
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
	   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">DESCRIPTION</td>
	   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">SUPPLIER</td>
	   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">ICNO</td>
	   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">BANK</td>
	   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">BANK ACCOUNT NO</td>
	   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">COMPANY NO</td>
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
          <td align="left"  title="PRODUCTCODE" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsVehicle("CUSTOMERCODE")%>&nbsp;</td>
	      <td align="left"  title="PRODUCTDESCRIPTION" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsVehicle("DESCRIPTION")%>&nbsp;</td>
	      <td align="left"  title="CROPCODE" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsVehicle("SUPPLIER")%>&nbsp;</td>
	      <td align="left"  title="UNITOFMEASURE" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsVehicle("ICNO")%>&nbsp;</td>
	      <td align="right"  title="FORMULA" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsVehicle("BANK")%>&nbsp;</td>
	      <td align="right"  title="FORMULA" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsVehicle("BANKACCOUNTNO")%>&nbsp;</td>
	      <td align="right"  title="FORMULA" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsVehicle("COMPANYNO")%>&nbsp;</td>
		  
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