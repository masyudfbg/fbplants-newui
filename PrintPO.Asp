<%'Option Explicit 
%>
<%'On Error GOTO 0 
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
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"
%>
 <%
 Dim txtNoPO,txtDatePO,txtRemarks,q,nCountofField,no,i
 Dim Remarks

   Dim rsPO,intRecordsPerPage,cColor,txtSupplier
   set rsPO = server.CreateObject("Adodb.recordset")
   q = " select fcba" & _
	   "        , pocode" & _
	   "        , podate" & _
	   "        , suppliercode" & _
	   "        , supplier_api.Get_Supp_Name(suppliercode, fcba) suppliername" & _
	   "        , remarks" & _
	   "        , rowstateid" & _
	   " from purchaseorder" & _
	   " where pocode = '" & request("fcCode") & "'" & _
	   " and fcba = '" & session("fcBA") & "'"
	   
   'response.Write q	
   Set rsPO = Server.CreateObject("ADODB.Recordset")
   rsPO.Open q, connstr
   if not rsPO.eof then
      txtNoPO = rsPO("POCODE")
	  txtDatePO = rsPO("PODATE")
	  txtRemarks = rsPO("REMARKS")
	  txtHeaderRemarks = rsPO("REMARKS")
	  txtSupplier = rsPO("suppliername")
   End if
   rsPO.close
   set rsPO = Nothing
   %>
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
	<tr><td width="100">Report</td><td  width="300"><b><FONT color="#000099">PURCHASE ORDER</FONT></td></tr>
	<tr><td width="100">Supplier</td><td  width="300"><b><FONT color="#000099"><%=txtSupplier%></FONT></td></tr>
	<tr><td width="100">PO Code</td><td  width="300"><b><FONT color="#000099"><%=request("fcCode")%></FONT></td></tr>	
	
 </table>

   <%
   Dim rsPurchaseOrder,nQty,nAmount
   set rsPurchaseOrder = server.CreateObject("Adodb.recordset")
   q="select PO.ITEMCODE,s.fcname,PO.OTHERITEMDESC,nvl(sum(po.Quantity),0) as Quantity,nvl(sum(po.Amount),0) as Amount from PURCHASEORDERDETAILS PO,stock s " & _
     " where (s.fccode = PO.itemcode and s.fcba=po.fcba)" &_
	 " and PO.PoCode='" & Request("fcCode") & "' and " & _
     " PO.fcBA='" & session("fcBA") & "' group by PO.ITEMCODE,s.fcname,PO.OTHERITEMDESC"
   'response.Write q	
   Set rsPurchaseOrder = Server.CreateObject("ADODB.Recordset")
   rsPurchaseOrder.CursorType = 3
   rsPurchaseOrder.CursorLocation = 3
   rsPurchaseOrder.Open q, connstr%>
   <br><br>
<table Cellpadding="2" bgcolor="#F7F7F7" cellspacing="0">
   <tr >
   <td style="<%=cmdstyleLeft%>;border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF" align="right"><font face="Verdana" size="1">NO</td>
   <td style="<%=cmdstyleTOP%>" align="center"><font face="Verdana" size="1">Stock<br>Code</td>
   <td style="<%=cmdstyleTOP%>" align="center" width="276"><font face="Verdana" size="1">Stock Description</td>
   <td style="<%=cmdstyleTOP%>" align="center" width="176"><font face="Verdana" size="1">Remarks</td>
   <td style="<%=cmdstyleTOP%>" align="center"><font face="Verdana" size="1">Quantity</td>
   <td style="<%=cmdstyleTOP%>" align="center"><font face="Verdana" size="1">Amount</td>
   </tr>
   <%
   if Not rsPurchaseOrder.eof then
      nCountofField=rsPurchaseOrder.fields.count-1
	  no=1
	  nQty=0
	  nAmount=0
	  While Not rsPurchaseOrder.EOF

		  nQty = nQty+cdbl(rsPurchaseOrder("Quantity"))
		  nAmount = nAmount+cdbl(rsPurchaseOrder("Amount"))
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr >
    	  <td align="left" style="<%=cmdstyleLeft%>"><font face="Verdana" size="1"><%=NO%>&nbsp;</td>
	      <td align="left"  style="<%=cmdstyle%>" title="PO.ITEMCODE" valign="top"><font face="Verdana" size="1"><%=rsPurchaseOrder("ITEMCODE")%>&nbsp;</td>
	      <td align="left"  style="<%=cmdstyle%>" title="s.fcname" valign="top"><font face="Verdana" size="1"><%=rsPurchaseOrder("fcname")%>&nbsp;</td>
	      <td align="left"  style="<%=cmdstyle%>" title="PO.OTHERITEMDESC" valign="top"><font face="Verdana" size="1"><%=rsPurchaseOrder("OTHERITEMDESC")%>&nbsp;</td>
	      <td align="Right"  style="<%=cmdstyle%>" title="po.Quantity" valign="top"><font face="Verdana" size="1"><%=FormatNumber(rsPurchaseOrder("Quantity"),2)%>&nbsp;</td>
	      <td align="Right"  style="<%=cmdstyle%>" title="po.Amount" valign="top"><font face="Verdana" size="1"><%=FormatNumber(rsPurchaseOrder("Amount"),2)%>&nbsp;</td>
		  </tr>
		  <%
		  rsPurchaseOrder.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend%>
	<tr>
	<td align="Center" style="<%=cmdstyleLeft%>" colspan="4"> <font face="Verdana" size="1">T o t a l</td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(nQty,2)%>&nbsp;</td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(nAmount,2)%>&nbsp;</td>
	</tr>
<%
   else
      response.Write("Data is empty")	  
   end if 
   rsPurchaseOrder.Close
   set rsPurchaseOrder = Nothing%>
   </TABLE>
<!--#include file="Bottom_Report.asp"-->