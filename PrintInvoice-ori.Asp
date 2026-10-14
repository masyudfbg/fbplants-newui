<%'Option Explicit
%>
<%'On Error GOTO CmdErrorlog 
%>
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
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"%>
 <%
 Dim txtNoReceved,txtDateReceived,txtRemarks,q,nCountofField,no,i
 Dim Remarks

   Dim rsReceived,intRecordsPerPage,cColor,txtSupplier
   set rsReceived = server.CreateObject("Adodb.recordset")
   q="select * from Invoice  " & _
     " where INVOICECODE='" & request("fcCode") & "' and " & _
     " fcBA='" & session("fcBA") & "'"
   'response.Write q	
   Set rsReceived = Server.CreateObject("ADODB.Recordset")
   rsReceived.Open q, connstr
   if not rsReceived.eof then
      txtNoReceved = rsReceived("INVOICECODE")
	  txtDateReceived = rsReceived("INVOICEDATE")
	  txtSupplier =cmdFindTitle("Supplier","fcCode",rsReceived("SUPPLIERCODE"))
	  txtHeaderRemarks="From Invoice Date : " & rsReceived("INVOICEDATE")
   End if
   rsReceived.close
   set rsReceived = Nothing%>
 <table width="700">
  <tr> 
    <td rowspan="7" width="100">
	<% if len(trim(PROFIL("REPORT_IMAGE"))) > 0 then %>
		<img src="IMAGES/<%=PROFIL("REPORT_IMAGE")%>" width="210" height="80">
	<% end if %>
	</td>
  </tr>
	<tr><td width="100">Company</td><td  width="300"><b><FONT color="#000099"><%=Session("COMPANY")%></FONT></td></tr>
	<tr><td width="100">Business Unit</td><td  width="300"><b><FONT color="#000099"><%=Session("BA")%> / <%=Session("FCBA")%></FONT></td></tr>
	<tr><td width="100">Report</td><td  width="300"><b><FONT color="#000099">INVOICE NOTE</FONT></td></tr>
 	<tr><td width="100">Supplier</td><td  width="300"><b><FONT color="#000099"><%=txtSupplier%></FONT></td></tr>
	<tr><td width="100">Invoice Code</td><td  width="300"><b><FONT color="#000099"><%=request("fcCode")%></FONT></td></tr>
</table>

   <%
   Dim rsPurchaseReceive,nQty,nAmount
   set rsPurchaseReceive = server.CreateObject("Adodb.recordset")
   q="select inv.INVOICEDATE,invd.pocode,invd.ITEMCODE,s.FCNAME as StockDesc,nvl(sum(invd.QUANTITY),0) as QUANTITY," & _
     "nvl(sum(invd.AMOUNT),0) as Amount From invoice inv,invoicedetail invd,stock s,supplier sp where" & _
     "(inv.INVOICECODE = invd.INVOICECODE and inv.fcba=invd.fcba) AND (INVD.ITEMCODE = S.FCCODE and invd.fcba=s.fcba) AND (INV.SUPPLIERCODE = SP.FCCODE and inv.fcba=sp.fcba) and " & _
     " inv.INVOICECODE='" & request("fcCode") &  "' and " & _
	 " inv.fcBA='" & session("fcBA") & "'" & _
	 " group by inv.INVOICEDATE ,invd.pocode,invd.ITEMCODE,s.FCNAME"
   'response.Write q	
   Set rsPurchaseReceive = Server.CreateObject("ADODB.Recordset")
   rsPurchaseReceive.CursorType = 3
   rsPurchaseReceive.CursorLocation = 3
   rsPurchaseReceive.Open q, connstr%>
   <br><br>
<table Cellpadding="2" bgcolor="#F7F7F7" cellspacing="0">
   <tr >
   <td style="<%=cmdstyleLeft%>;border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF" align="right"><font face="Verdana" size="1">No</FONT></td>
   <td  style="<%=cmdstyleTOP%>"align="center"><font face="Verdana" size="1">Date</FONT></td>
   <td  style="<%=cmdstyleTOP%>"align="center"><font face="Verdana" size="1">PO<br>Code</FONT></td>
   <td  style="<%=cmdstyleTOP%>"align="center"><font face="Verdana" size="1">Stock<br>Code</FONT></td>
   <td  style="<%=cmdstyleTOP%>"align="center"><font face="Verdana" size="1">Stock<br>Description</FONT></td>
   <td  style="<%=cmdstyleTOP%>"align="center"><font face="Verdana" size="1">Quantity</FONT></td>
   <td  style="<%=cmdstyleTOP%>"align="center"><font face="Verdana" size="1">Amount</FONT></td>
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
		  <tr >
		  <td  style="<%=cmdstyleLeft%>" align="left" valign="top"><font face="Verdana" size="1"><%=NO%>&nbsp;</td>
	      <td  style="<%=cmdstyle%>"align="left" title="INVOICEDATE" valign="top"><font face="Verdana" size="1"><%=rsPurchaseReceive("INVOICEDATE")%>&nbsp;</td>
	      <td  style="<%=cmdstyle%>"align="left" title="ITEMCODE" valign="top"><font face="Verdana" size="1"><%=rsPurchaseReceive("pocode")%>&nbsp;</td>
	      <td  style="<%=cmdstyle%>"align="left" title="ITEMCODE" valign="top"><font face="Verdana" size="1"><%=rsPurchaseReceive("ITEMCODE")%>&nbsp;</td>
	      <td  style="<%=cmdstyle%>"align="left" title="StockDesc" valign="top"><font face="Verdana" size="1"><%=rsPurchaseReceive("StockDesc")%>&nbsp;</td>
	      <td  style="<%=cmdstyle%>"align="Right" title="QUANTITY" valign="top"><font face="Verdana" size="1"><%=FormatNumber(rsPurchaseReceive("QUANTITY"),2)%>&nbsp;</td>
	      <td  style="<%=cmdstyle%>"align="Right" title="Amount" valign="top"><font face="Verdana" size="1"><%=FormatNumber(rsPurchaseReceive("Amount"),2)%>&nbsp;</td>
		  </tr>
		  <%
		  rsPurchaseReceive.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend%>
	  <tr>
	  <td colspan="5" style="<%=cmdstyleLeft%>"align="center"><font face="Verdana" size="1">T o t a l</td>
	  <td style="<%=cmdstyle%>"align="right"><font face="Verdana" size="1"><%=CheckValue(nQty)%>&nbsp;</td>
	  <td style="<%=cmdstyle%>"align="right"><font face="Verdana" size="1"><%=CheckValue(nAmount)%>&nbsp;</td>
	  </tr><%
   else
      response.Write("Data is empty")	  
   end if 
   rsPurchaseReceive.Close
   set rsPurchaseReceive = Nothing%>
   </TABLE>
 <!--#include file="Bottom_Report.asp"-->