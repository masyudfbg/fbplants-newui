<%'Option Explicit%>
<%' On Error GOTO CmdErrorlog %>
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


   		q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & session("fcBA") & "' "
			' response.write q
			set rsCat = connStr.Execute(q)
			if not rsCat.eof then
			   fcbaname = rsCat("fcName")
			   CompanyName=rsCat("FCCOMPANYNAME")
			   CompanyCode=rsCat("fccompanycode")
			   MgrName = rsCat("FCMANAGER")
			   KTUName = rsCat("FCKTU")
			   CropName=rsCat("FCCROP")
			   FcbaType=rsCat("fcType")
			end if



   Dim rsSIV,intRecordsPerPage,cColor,txtStore
   set rsSIV = server.CreateObject("Adodb.recordset")
   q="select * from WriteOff  " & _
     " where WriteOffCode='" & request("fcCode") & "' and " & _
     " fcBA='" & session("fcBA") & "'"
   'response.Write q	
   Set rsSIV = Server.CreateObject("ADODB.Recordset")
   rsSIV.Open q, connstr
   if not rsSIV.eof then
      txtNoPO = rsSIV("WriteOffCODE")
	  txtDatePO = rsSIV("WODATE")
	  txtStore =cmdFindTitle("Store","fcCode",rsSIV("STORECODE"))
   End if
   rsSIV.close
   set rsSIV = Nothing%>
 <table width="700">
  <tr> 
    <td rowspan="6" width="100">
    <%if trim(CompanyCode) = "PT.ARJ" THEN%>
      <img src="css/login/images/logoarj.jpeg" width="210" height="80">
    <%end if%>
    <%if trim(CompanyCode) = "PT.BRS" THEN%>
      <img src="css/login/images/logobrs.jpg" width="210" height="80">
    <%end if%>
    <%if trim(CompanyCode) = "PT.BIA" THEN%>
      <img src="css/login/images/logobia.jpeg" width="210" height="80">
    <%end if%>
    <%if trim(CompanyCode) = "PT.KAA" THEN%>
      <img src="css/login/images/logokaa.jpg" width="210" height="80">
    <%end if%>
    <%if trim(CompanyCode) = "PT.POM" THEN%>
      <img src="css/login/images/logokaa.jpg" width="210" height="80">
    <%end if%>
    <%if trim(CompanyCode) = "PT.CPP" THEN%>
      <img src="css/login/images/logocpp.png" width="210" height="80">
    <%end if%>
    <%if trim(CompanyCode) = "PT.GAS" THEN%>
      <img src="css/login/images/logogas.jpeg" width="210" height="80">
    <%end if%>
    <%if trim(CompanyCode) = "PT.KSR" THEN%>
      <img src="css/login/images/logoksr.jpg" width="210" height="80">
    <%end if%>
    <%if trim(CompanyCode) = "PT.BRP" THEN%>
      <img src="css/login/images/logobrp.jpg" width="210" height="80">
    <%end if%>
    <%if trim(CompanyCode) = "PT.CAE" THEN%>
      <img src="css/login/images/logocae.jpg" width="210" height="80">
    <%end if%>
    </td>
  </tr>
	<tr><td width="100">Company</td><td  width="300"><b><FONT color="#000099"><%=Session("COMPANY")%></FONT></td></tr>
	<tr><td width="100">Business Unit</td><td  width="300"><b><FONT color="#000099"><%=Session("BA")%> / <%=Session("FCBA")%></FONT></td></tr>
	<tr><td width="100">Report</td><td  width="300"><b><FONT color="#000099">WRITE OFF</FONT></td></tr>
	<tr><td width="100">WO No</td><td  width="300"><b><FONT color="#000099"><%= request("fcCode")%></FONT></td></tr>
	<tr><td width="100">Date</td><td  width="300"><b><FONT color="#000099"><%=request("WODate")%></FONT></td></tr>
 </table>
   <%
   Dim rsPurchaseOrder,nQty,nAmount
   set rsPurchaseOrder = server.CreateObject("Adodb.recordset")
   q=" select wo.WriteOffCODE, wod.JOBCODE,Job.FCNAME as JobDesc,wod.STOCKCODE,S.FCNAME as StockDesc, wod.LOCATIONCODE, wod.LOCATIONTYPE, " & _
     " nvl(sum(wod.QUANTITY),0) as QUANTITY, nvl(sum(wod.PRICE),0) as PRICE, nvl(sum(wod.AMOUNT),0) as AMOUNT" & _
     " from WriteOff WO,WriteOffDetails WOD,job,stock s " & _
     " where (wo.WRITEOFFCODE = wod.WRITEOFFCODE and wo.fcba = wod.fcba) and" & _
     " (wod.jobcode = job.FCCODE and wod.FCBA = job.fcba) and" & _
	 " (wod.STOCKCODE = s.FCCODE and wod.FCBA = s.FCBA) and " & _
	 " wo.WriteOffCode='" & Request("fcCode") & "' and " & _
     " wo.fcBA='" & session("fcBA") & _
	 "' group by wo.WriteOffCODE, wod.JOBCODE,Job.FCNAME ,wod.STOCKCODE,S.FCNAME , wod.LOCATIONCODE, wod.LOCATIONTYPE"
   'response.Write q
   'response.Flush()	
   Set rsPurchaseOrder = Server.CreateObject("ADODB.Recordset")
   rsPurchaseOrder.CursorType = 3
   rsPurchaseOrder.CursorLocation = 3
   rsPurchaseOrder.Open q, connstr%>
   <br><br>
<table Cellpadding="2" bgcolor="#F7F7F7" cellspacing="0">
   <tr >
	    <td style="<%=cmdstyleLeft%>;border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF" align="right"><font face="Verdana" size="1">No</td>
   <td style="<%=cmdstyleTOP%>" align="center"><font face="Verdana" size="1">Code</td></td>
   <td style="<%=cmdstyleTOP%>"align="center"><font face="Verdana" size="1">Job<br>Code</td>
   <td style="<%=cmdstyleTOP%>"align="center"><font face="Verdana" size="1">Job<br>Description</td>
   <td style="<%=cmdstyleTOP%>"align="center"><font face="Verdana" size="1">Stock<br>Code</td>
   <td style="<%=cmdstyleTOP%>"align="center"><font face="Verdana" size="1">Stock<br>Description</td>
   <td style="<%=cmdstyleTOP%>"align="center"><font face="Verdana" size="1">Target<br>Type</td>
   <td style="<%=cmdstyleTOP%>"align="center"><font face="Verdana" size="1">Target<br>Code</td>
   <td style="<%=cmdstyleTOP%>"align="center"><font face="Verdana" size="1">Quantity</td>
   <td style="<%=cmdstyleTOP%>"align="center"><font face="Verdana" size="1">Amount</td>
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
		  <td align="left" style="<%=cmdstyleLeft%>"valign="top"><font face="Verdana" size="1"><%=NO%>&nbsp;</td>
          <td align="left" style="<%=cmdstyle%>" title="PO.ITEMCODE" valign="top"><font face="Verdana" size="1"><%=rsPurchaseOrder("WriteOffCode")%>&nbsp;</td>
	      <td align="left" style="<%=cmdstyle%>" title="s.fcname" valign="top"><font face="Verdana" size="1"><%=rsPurchaseOrder("JOBCODE")%>&nbsp;</td>
	      <td align="left" style="<%=cmdstyle%>" title="PO.OTHERITEMDESC" valign="top"><font face="Verdana" size="1"><%=rsPurchaseOrder("JobDesc")%>&nbsp;</td>
	      <td align="left" style="<%=cmdstyle%>" title="PO.OTHERITEMDESC" valign="top"><font face="Verdana" size="1"><%=rsPurchaseOrder("STOCKCODE")%>&nbsp;</td>
	      <td align="left" style="<%=cmdstyle%>" title="PO.OTHERITEMDESC" valign="top"><font face="Verdana" size="1"><%=rsPurchaseOrder("StockDesc")%>&nbsp;</td>
	      <td align="left" style="<%=cmdstyle%>" title="PO.OTHERITEMDESC" valign="top"><font face="Verdana" size="1"><%=rsPurchaseOrder("LOCATIONTYPE")%>&nbsp;</td>
	      <td align="left" style="<%=cmdstyle%>" title="PO.OTHERITEMDESC" valign="top"><font face="Verdana" size="1"><%=rsPurchaseOrder("LOCATIONCODE")%>&nbsp;</td>
	      <td align="Right" style="<%=cmdstyle%>" title="po.Quantity" valign="top"><font face="Verdana" size="1"><%=FormatNumber(rsPurchaseOrder("Quantity"),2)%>&nbsp;</td>
	      <td align="Right" style="<%=cmdstyle%>" title="po.Amount" valign="top"><font face="Verdana" size="1"><%=FormatNumber(rsPurchaseOrder("AMOUNT"),2)%>&nbsp;</td>
		  </tr>
		  <%
		  rsPurchaseOrder.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend%>
	  <tr >
	  <td colspan="8" style="<%=cmdstyleLeft%>" align="center"><font face="Verdana" size="1">T o t a l</td>
	  <td align="right"style="<%=cmdstyle%>" ><font face="Verdana" size="1"><%=FormatNumber(nQty,2)%>&nbsp;</td>
	  <td align="right"style="<%=cmdstyle%>" ><font face="Verdana" size="1"><%=FormatNumber(nAmount,2)%>&nbsp;</td>
	  </tr><%
   else
      response.Write("Data is empty")	  
   end if 
   rsPurchaseOrder.Close
   set rsPurchaseOrder = Nothing%>
   </TABLE>
<!--#include file="Bottom_Report.asp"-->