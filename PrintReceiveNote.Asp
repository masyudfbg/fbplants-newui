<%'Option Explicit%>
<%'On Error GOTO CmdErrorlog %>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->

<%response.buffer=True%>
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"
 Dim txtNoReceved,txtDateReceived,txtRemarks,q,nCountofField,no,i
 Dim Remarks

   Dim rsReceived,intRecordsPerPage,cColor,txtSupplier
   set rsReceived = server.CreateObject("Adodb.recordset")
   q="select * from ReceiveNote  " & _
     " where RECEIVENOTECODE='" & request("fcCode") & "' and " & _
     " fcBA='" & session("fcBA") & "'"
   'response.Write q	
   Set rsReceived = Server.CreateObject("ADODB.Recordset")
   rsReceived.Open q, connstr
   if not rsReceived.eof then
      txtNoReceved = rsReceived("RECEIVENOTECODE")
	  txtDateReceived = rsReceived("RNDATE")
	  txtSupplier =cmdFindTitle("Supplier","fcCode",rsReceived("SUPPLIERCODE"))
	  txtHeaderRemarks="From PO Code : " & rsReceived("POCODE")
   End if
   rsReceived.close
   set rsReceived = Nothing%>
 <table width="700">
  <tr> 
    <td rowspan="7" width="100"><img src="IMAGES/latar.jpg" width="210" height="80"></td>
  </tr>
	<tr><td width="100">Company</td><td  width="300"><b><FONT color="#000099"><%=Session("COMPANY")%></FONT></td></tr>
	<tr><td width="100">Business Unit</td><td  width="300"><b><FONT color="#000099"><%=Session("BA")%> / <%=Session("FCBA")%></FONT></td></tr>
	<tr><td width="100">Report</td><td  width="300"><b><FONT color="#000099">RECEIVE NOTE</FONT></td></tr>
	<tr><td width="100">Supplier</td><td  width="300"><b><FONT color="#000099"><%=txtSupplier%></FONT></td></tr>
	<tr><td width="100">GRN Code</td><td  width="300"><b><FONT color="#000099"><%=request("fcCode")%></FONT></td></tr>
	<tr><td width="100">Date</td><td  width="300"><b><FONT color="#000099"><%=request("DateCreated")%></FONT></td></tr>
 </table>

   <%
   Dim rsPurchaseReceive,nQty,nAmount
   set rsPurchaseReceive = server.CreateObject("Adodb.recordset")
   q="select p.RECEIVENOTECODE, p.LOCATIONTYPE, p.LOCATIONCODE, p.PURCHASEITEMCODE, s.fcname,p.JOBCODE,j.fcname as JobDesc, " & _
     "nvl(sum(p.QUANTITY),0) as Quantity, nvl(sum(p.AMOUNTRECEIVED),0) as AMOUNT" & _
     " from RECEIVENOTEDETAIL p ,Stock s,Job j " & _
        " where RECEIVENOTECODE='" & txtNoReceved & "' and (p.PURCHASEITEMCODE=s.fccode AND p.fcba=s.fcba) and (j.fccode=p.jobcode and j.fcba=p.fcba) AND" & _
	    " p.fcBA='" & session("fcBA") & "' group by p.RECEIVENOTECODE, p.LOCATIONTYPE, p.LOCATIONCODE, p.PURCHASEITEMCODE, s.fcname,p.JOBCODE,j.fcname"
   'response.Write q	
   Set rsPurchaseReceive = Server.CreateObject("ADODB.Recordset")
   rsPurchaseReceive.CursorType = 3
   rsPurchaseReceive.CursorLocation = 3
   rsPurchaseReceive.Open q, connstr%>
   <br><br>
  <table Cellpadding="2" bgcolor="#F7F7F7" cellspacing="0">
   <tr>
   <td style="<%=cmdstyleLeft%>;border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF" align="right"><font face="Verdana" size="1">No.</font></td>
   <td style="<%=cmdstyleTOP%>" align="center"><font face="Verdana" size="1">Target<br>Type</font></td>
   <td style="<%=cmdstyleTOP%>" align="center"><font face="Verdana" size="1">Target<br>Code</font></td>
   <td style="<%=cmdstyleTOP%>" align="center"><font face="Verdana" size="1">Stock<br>Code</font></td>
   <td style="<%=cmdstyleTOP%>" align="center"><font face="Verdana" size="1">Stock<br>Desc.</font></td>
   <td style="<%=cmdstyleTOP%>" align="center"><font face="Verdana" size="1">Job<br>Code</font></td>
   <td style="<%=cmdstyleTOP%>" align="center"><font face="Verdana" size="1">Job<br>Description</font></td>
   <td style="<%=cmdstyleTOP%>" align="center"><font face="Verdana" size="1">Quantity</font></td>
   <td style="<%=cmdstyleTOP%>" align="center"><font face="Verdana" size="1">Amount</font></td>
   </tr>
   <%
   if Not rsPurchaseReceive.eof then
      nCountofField=rsPurchaseReceive.fields.count-1
	  no=1
	  nQty=0
	  nAmount=0
	  While Not rsPurchaseReceive.EOF

		  nQty = nQty+CDBL(rsPurchaseReceive("QUANTITY"))
		  nAmount = nAmount+CDBL(rsPurchaseReceive("AMOUNT"))
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr>
		<td align="left" style="<%=cmdstyleLeft%>"><font face="Verdana" size="1"><%=NO%>&nbsp;</td>
	      <td align="left"  style="<%=cmdstyle%>" title="LOCATIONTYPE" valign="top"><font face="Verdana" size="1"><%=rsPurchaseReceive("LOCATIONTYPE")%></td>
	      <td align="left"  style="<%=cmdstyle%>" title="LOCATIONCODE" valign="top"><font face="Verdana" size="1"><%=rsPurchaseReceive("LOCATIONCODE")%></td>
	      <td align="left"  style="<%=cmdstyle%>" title="PURCHASEITEMCODE" valign="top"><font face="Verdana" size="1"><%=rsPurchaseReceive("PURCHASEITEMCODE")%></td>
	      <td align="left"  style="<%=cmdstyle%>" title="fcname" valign="top"><font face="Verdana" size="1"><%=rsPurchaseReceive("fcname")%></td>
	      <td align="left"  style="<%=cmdstyle%>" title="JOBCODE" valign="top"><font face="Verdana" size="1"><%=rsPurchaseReceive("JOBCODE")%></td>
	      <td align="left"  style="<%=cmdstyle%>" title="JOBCODE" valign="top"><font face="Verdana" size="1"><%=rsPurchaseReceive("JOBDESC")%></td>
	      <td align="Right"  style="<%=cmdstyle%>" title="QUANTITY" valign="top"><font face="Verdana" size="1"><%=FormatNumber(rsPurchaseReceive("QUANTITY"),2)%></td>
	      <td align="Right"  style="<%=cmdstyle%>" title="QTYRETURN" valign="top"><font face="Verdana" size="1"><%=FormatNumber(rsPurchaseReceive("AMOUNT"),2)%></td>
		  </tr>
		  <%
		  rsPurchaseReceive.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend%>
	  <tr>
	  <td colspan="7" align="center"  style="<%=cmdstyleLeft%>" ><font face="Verdana" size="1">T o t a l</td>
	  <td align="right" style="<%=cmdstyle%>" ><font face="Verdana" size="1"><%=FormatNumber((nQty),2)%></td>
	  <td align="right" style="<%=cmdstyle%>" ><font face="Verdana" size="1"><%=FormatNumber((nAmount),2)%></td>
	  </tr><%
   else
      response.Write("Data is empty")	  
   end if 
   rsPurchaseReceive.Close
   set rsPurchaseReceive = Nothing%>
   </TABLE>
<!--#include file="Bottom_Report.asp"-->