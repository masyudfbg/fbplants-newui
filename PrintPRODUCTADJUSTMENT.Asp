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
		<tr><td width="100">Report</td><td  width="300"><b><FONT color="#000099">PRODUCT ADJUSMENT</FONT></td></tr>
		<tr><td width="100">For Date</td><td  width="300"><b><FONT color="#000099"><%=FormatDateTime(DATE(),1)%></FONT></td></tr>
		<tr><td width="100">Project Code</td><td  width="300"><b><FONT color="#000099"><%=request("txt_ADJUSTMENTNO")%></FONT></td></tr>
	 </table>
	<BR><BR>


   <%
   Dim rsVehicle
   set rsVehicle = server.CreateObject("Adodb.recordset")  
 q="select ADJUSTMENTDATE, ADJUSTMENTNO, PSTORAGE_STORAGECODE, QUANTITY, FCENTRY, FCEDIT, FCIP, FCBA from PRODUCTADJUSTMENT where " & _
     " fcBA='" & session("fcBA") & "'"
	if  len(trim(request("Perioddate"))) > 0 then
		q = q & " and to_char(ADJUSTMENTDATE,'mm/yyyy')=to_char(to_date('" & request("Perioddate") & "','" & "MM/DD/YYYY" & "'),'mm/yyyy')"
	end if
	if  len(trim(request("txt_STORAGECODE"))) > 0 then
		q = q & " and PSTORAGE_STORAGECODE='" & request("txt_STORAGECODE") & "'"
	end if
	if  len(trim(request("txt_ADJUSTMENTNO"))) > 0 then
		q = q & " and ADJUSTMENTNO='" & request("txt_ADJUSTMENTNO") & "'"
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
	   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">DATE</td></td>
	   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">ADJUSTMENT NO</td>
	   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">STORAGE CODE</td>
	   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">QUANTITY</td>
   </tr>
   <%
   'Dim nMillweight,nEstate,nQuantity
   if Not rsVehicle.eof then
      nCountofField=rsVehicle.fields.count-1
	  no=1
	'	nMillweight=0
	'	nEstate=0
	'	nQuantity=0
	  While Not rsVehicle.EOF
		  'nMillweight = nMillweight+cdbl(rsVehicle("MILL_WEIGHT"))
		  'nEstate = nEstate+cdbl(rsVehicle("ESTATE_WEIGHT"))
		  'nQuantity = nQuantity+cdbl(rsVehicle("QUANTITY"))
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td align="left"  style="<%=cmdstyleLeft%>" valign="top"><font size="1"><%=NO%></font></td>
          <td align="left"  title="CHITNO" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsVehicle("ADJUSTMENTDATE")%></td>
	      <td align="left"  title="FIELDCODE" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsVehicle("ADJUSTMENTNO")%></td>
	      <td align="left"  title="CROPCODE" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsVehicle("PSTORAGE_STORAGECODE")%></td>
	      <td align="left"  title="CROPCODE" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsVehicle("QUANTITY")%></td>
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