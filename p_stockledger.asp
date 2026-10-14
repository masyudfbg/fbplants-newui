<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>ePlaS Reporting - Stock Ledger</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<script language="JavaScript" src='/js/x.js'></script>
<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
<script language="JavaScript">
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window);
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);
</script>
</head>

<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

v_txtyear=trim(request("txtyear"))
v_txtmonth=trim(request("txtmonth"))
v_txtstockgroup=trim(request("txtstockgroup"))
v_txtstock=trim(request("txtstock"))
v_txtstorecode=trim(request("txtstorecode"))

txtSQL =	" SELECT fcba," & _
			"        storecode," & _
			"        stockcode," & _
			"        stockname," & _
			"        uom," & _
			"        controljobcode," & _
			"        controljobname," & _
			"        stockgroup," & _
			"        openingbalanceqty," & _
			"        qtyreceived," & _
			"        QTYRECEIVEDTASK," & _
			"        QTYTRANSPORTTASK," & _
			"        qtywriteoff," & _
			"        qtyissued," & _
			"        qtysaldo," & _
			"        v_saldo," & _
			"        openingbalanceamt," & _
			"        amtreceived, " & _
			"		 AMTRECEIVEDTASK," & _
			"        AMTTRANSPORTTASK," & _
			"        amtwriteoff," & _
			"        amtissued," & _
			"        amtsaldo," & _
			"        amtperunit" & _
			"   FROM (  SELECT fcba," & _
			"                  storecode," & _
			"                  itemcode stockcode," & _
			"                  stock_api.get_stockname (itemcode, fcba) stockname," & _
			"                  stock_api.get_uom (itemcode, fcba) uom," & _
			"                  stockgroup_api.get_controljob ( stock_api.get_stockgroup (itemcode, fcba), fcba) controljobcode," & _
			"                  job_api.get_jobcode_name ( stockgroup_api.get_controljob (stock_api.get_stockgroup (itemcode, fcba), fcba), fcba) controljobname," & _
			"                  stock_api.get_stockgroup (itemcode, fcba) || ' - ' || stockgroup_api.Get_StockGrpName ( stock_api.get_stockgroup (itemcode, fcba), fcba) stockgroup," & _
			"                  NVL (SUM (NVL (openingbalanceqty, 0)), 0) openingbalanceqty," & _
			"                  NVL (SUM (NVL (qtyreceived, 0)), 0) qtyreceived," & _
			"                  NVL (SUM (NVL (QTYRECEIVEDTASK, 0)), 0) QTYRECEIVEDTASK," & _
			"                  NVL (SUM (NVL (QTYTRANSPORTTASK, 0)), 0) QTYTRANSPORTTASK," & _
			"                  NVL (SUM (NVL (qtywriteoff, 0)), 0) qtywriteoff," & _
			"                  NVL (SUM (NVL (qtyissued, 0)), 0) qtyissued," & _
			"                    NVL (SUM (NVL (openingbalanceqty, 0)), 0)" & _
			"                  + NVL (SUM (NVL (qtyreceived, 0)), 0)" & _
			"                  + NVL (SUM (NVL (QTYRECEIVEDTASK, 0)), 0)" & _
			"                  + NVL (SUM (NVL (qtywriteoff, 0)), 0)" & _
			"                  + (NVL (SUM (NVL (QTYTRANSPORTTASK, 0)), 0) * -1)" & _
			"                  + (NVL (SUM (NVL (qtyissued, 0)), 0) * -1)" & _
			"                     qtysaldo," & _
			"                    NVL (SUM (NVL (openingbalanceqty, 0)), 0)" & _
			"                  + NVL (SUM (NVL (qtyreceived, 0)), 0)" & _
			"                  + NVL (SUM (NVL (QTYRECEIVEDTASK, 0)), 0)" & _
			"                  + NVL (SUM (NVL (qtywriteoff, 0)), 0)" & _ 
			"                     v_saldo," & _
			"                  NVL (SUM (NVL (openingbalanceamt, 0)), 0) openingbalanceamt," & _
			"                  NVL (SUM (NVL (amtreceived, 0)), 0) amtreceived," & _
			"                  NVL (SUM (NVL (AMTRECEIVEDTASK, 0)), 0) AMTRECEIVEDTASK," & _
			"                  NVL (SUM (NVL (AMTTRANSPORTTASK, 0)), 0) AMTTRANSPORTTASK," & _
			"                  NVL (SUM (NVL (amtwriteoff, 0)), 0) amtwriteoff," & _
			"                  NVL (SUM (NVL (amtissued, 0)), 0) amtissued," & _
			"                    NVL (SUM (NVL (openingbalanceamt, 0)), 0)" & _
			"                  + NVL (SUM (NVL (amtreceived, 0)), 0)" & _
			"                  + NVL (SUM (NVL (AMTRECEIVEDTASK, 0)), 0)" & _
			"                  + NVL (SUM (NVL (amtwriteoff, 0)), 0)" & _
			"                  + (NVL (SUM (NVL (AMTTRANSPORTTASK, 0)), 0) * -1)" & _
			"                  + (NVL (SUM (NVL (amtissued, 0)), 0) * -1)" & _
			"                     amtsaldo," & _
			"                  ROUND (NVL (SUM (NVL (amtperunit, 0)), 0), 2) amtperunit" & _
			"             FROM stockledger" & _
			"            WHERE     fcba = '"& session("fcba") &"'" & _
			"                  AND itemcode NOT IN ('002334')"  
 
if  len(trim(v_txtyear)) > 0 then
    txtSQL = txtSQL & " and year = " & v_txtyear
	txtSelect = txtSelect & " Year " & v_txtyear
end if
if  len(trim(v_txtmonth)) > 0 then
    txtSQL =txtSQL &  " and month = " & v_txtmonth
	txtSelect = txtSelect & " Month " & v_txtmonth
end if
if  len(trim(v_txtstorecode)) > 0 then
    txtSQL = txtSQL & " and storecode = '" & v_txtstorecode & "'"
	txtSelect = txtSelect & " Store Code " & v_txtstorecode
end if
if  len(trim(v_txtstockgroup)) > 0 then
    txtSQL = txtSQL & " and itemcode in (select fccode from stock where stockgroupcode = " & v_txtstockgroup & ")"
	txtSelect = txtSelect & " Stock Group " & v_txtstockgroup
end if
if  len(trim(v_txtstock)) > 0 then
    txtSQL =txtSQL &  " and itemcode = '" & v_txtstock & "'"
	txtSelect = txtSelect & " Stock Code " & v_txtstock
end if 
txtsql = txtsql &	" group by fcba" & _
					"          , storecode" & _
					"          , itemcode"
txtsql = txtsql &	" order by fcba" & _
					"          , storecode" & _
					"          , stockgroup_api.get_controljob(stock_api.get_stockgroup(itemcode, fcba), fcba)" & _
					"          , stock_api.get_stockgroup(itemcode, fcba)" & _
					"          , itemcode)"
txtsql = txtsql & "  WHERE fcba = '"& session("fcba") &"' AND v_saldo <> 0 OR qtywriteoff > 0 OR amtwriteoff > 0" 

' response.Write(txtsql)
response.Flush()
%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<body>
		<div id="wrapper">
		<!--#include file="print_excel.asp"-->
         <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="20" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=Session("COMPANY")%></h4>
					<h5><%=Session("BA")%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>STOCK LEDGER</h2>
				</div>
			</div>
			  </th>
			</tr>
	<!--<div class="printhide">-->
						<tr>
							<th colspan="20" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
						</tr>
						<!--</div>-->
	<tr>
	<%	
	Set rsCat = connStr.Execute(txtSQL)
	if Not rsCat.eof then
	   nCountofField=rsCat.fields.count-1
	end if
	NO=1%>
	</tr>
	<TR>
		<th rowspan="2">Store Code</th>
		<th rowspan="2">Stock Group</th>
		<th rowspan="2">Stock Code</th>
		<th rowspan="2">Description</th>
		<th rowspan="2">UOM</th>
		<Th colspan="7">Quantity</Th>
		<Th colspan="7">Amount (Rp)</Th>
		<th rowspan="2">Cost Per Unit</th>
	</TR>
	<tr>
		<th>Opening</th>
		<th>Received</th>
		<th>Received Task</th>
		<th>WriteOff</th>
		<th>Issued</th>
		<th>Transport Task</th>
		<th>Saldo</th>
		<th>Opening</th>
		<th>Received</th>
		<th>Received Task</th>
		<th>Write Off</th>
		<th>Issued</th>
		<th>Transport Task</th>
		<th>Saldo</th>
	</tr>
  </thead>
  <tbody>
	<%
	While Not rsCat.EOF
		Currcontroljobcode = rsCat("controljobcode")
        Currcontroljobname = rsCat("controljobname")
		If Currcontroljobname <> Prevcontroljobname Then
%>
<%If Len(Trim(Prevcontroljobname)) = 0 Then%>
		<tr>
			<td colspan="20" align="left" style="padding-left:10px;" bgcolor="#F4F4A0"><font face="Verdana" size="1"><strong><%=Currcontroljobname%></strong></font></td>
		</tr>
<%
		subopeningbalanceamt = 0
		subamtreceived = 0
		subamtwriteoff = 0
		subamtissued = 0
		subamtsaldo = 0
		subamttransporttask=0
		subamtreceivedtask=0
%>
<%Else%>
		<tr bgcolor="#9CB770">
			<td align="left" colspan="12" style="padding-left:10px;"><font face="Verdana" size="1"><strong>Sub Total <%=PCase(Prevcontroljobname)%></strong></font></td>
			<td align="Right"><font face="Verdana" size="1"><strong><%=FormatNumber(subopeningbalanceamt, 2)%></strong></font></td>
			<td align="Right"><font face="Verdana" size="1"><strong><%=FormatNumber(subamtreceived, 2)%></strong></font></td>
			<td align="Right"><font face="Verdana" size="1"><strong><%=FormatNumber(subamtreceivedtask, 2)%></strong></font></td>
			<td align="Right"><font face="Verdana" size="1"><strong><%=FormatNumber(subamtwriteoff, 2)%></strong></font></td>
			<td align="Right"><font face="Verdana" size="1"><strong><%=FormatNumber(subamtissued, 2)%></strong></font></td>
			<td align="Right"><font face="Verdana" size="1"><strong><%=FormatNumber(subamttransporttask, 2)%></strong></font></td>
			<td align="Right"><font face="Verdana" size="1"><strong><%=FormatNumber(subamtsaldo, 2)%></strong></font></td>
			<td align="Right"><font face="Verdana" size="1"><strong>&nbsp;</strong></font></td>
		</tr>
		<tr>
			<td colspan="20" align="left" style="padding-left:10px;" bgcolor="#F4F4A0"><font face="Verdana" size="1"><strong><%=Currcontroljobname%></strong></font></td>
		</tr>
<%
subopeningbalanceamt = 0
subamtreceived = 0
subamtwriteoff = 0
subamtissued = 0
subamtsaldo = 0
subamttransporttask=0
subamtreceivedtask=0
%>
<%End If%>
		<tr>
		      <td align="left"><font face="Verdana" size="1"><%=rsCat("storecode")%></font></td>
			  <td align="left"><font face="Verdana" size="1"><%=rsCat("stockgroup")%></font></td>
		      <td align="left"><font face="Verdana" size="1"><%=rsCat("stockcode")%></font></td>
		      <td align="left"><font face="Verdana" size="1"><%=rsCat("stockname")%></font></td>
			  <td align="left"><font face="Verdana" size="1"><%=rsCat("uom")%></font></td>
		      <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("openingbalanceqty"), 2)%></font></td>
		      <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("qtyreceived"), 2)%></font></td>
			  <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("QTYRECEIVEDTASK"), 2)%></font></td>
		      <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("qtywriteoff"), 2)%></font></td>
		      <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("qtyissued"), 2)%></font></td>
			   <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("QTYTRANSPORTTASK"), 2)%></font></td>
		      <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("qtysaldo"), 2)%></font></td>
		      <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("openingbalanceamt"), 2)%></font></td>
		      <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("amtreceived"), 2)%></font></td>
			  <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("AMTRECEIVEDTASK"), 2)%></font></td>
		      <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("amtwriteoff"), 2)%></font></td>		      
		      <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("amtissued"), 2)%></font></td>
			   <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("AMTTRANSPORTTASK"), 2)%></font></td>
		      <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("amtsaldo"), 2)%></font></td>
			  <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("amtperunit"), 2)%></font></td>
		</tr>
<%
		  subopeningbalanceamt = subopeningbalanceamt + cdbl(rsCat("openingbalanceamt"))
		  subamtreceived = subamtreceived + cdbl(rsCat("amtreceived"))
		  subamtwriteoff = subamtwriteoff + cdbl(rsCat("amtwriteoff"))
		  subamtissued = subamtissued + cdbl(rsCat("amtissued"))
          subamtsaldo = subamtsaldo + cdbl(rsCat("amtsaldo"))
		  subamttransporttask= subamttransporttask + cdbl(rsCat("AMTTRANSPORTTASK"))
		  subamtreceivedtask= subamtreceivedtask + cdbl(rsCat("AMTRECEIVEDTASK"))
%>
<%Else%>
		<tr>
		      <td align="left"><font face="Verdana" size="1"><%=rsCat("storecode")%></font></td>
			  <td align="left"><font face="Verdana" size="1"><%=rsCat("stockgroup")%></font></td>
		      <td align="left"><font face="Verdana" size="1"><%=rsCat("stockcode")%></font></td>
		      <td align="left"><font face="Verdana" size="1"><%=rsCat("stockname")%></font></td>
			  <td align="left"><font face="Verdana" size="1"><%=rsCat("uom")%></font></td>
		      <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("openingbalanceqty"), 2)%></font></td>
		      <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("qtyreceived"), 2)%></font></td>
			   <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("QTYRECEIVEDTASK"), 2)%></font></td>
		      <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("qtywriteoff"), 2)%></font></td>
		      <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("qtyissued"), 2)%></font></td>
			  <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("QTYTRANSPORTTASK"), 2)%></font></td>
		      <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("qtysaldo"), 2)%></font></td>
		      <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("openingbalanceamt"), 2)%></font></td>
		      <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("amtreceived"), 2)%></font></td>
			  <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("AMTRECEIVEDTASK"), 2)%></font></td>
		      <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("amtwriteoff"), 2)%></font></td>		      
		      <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("amtissued"), 2)%></font></td>
			   <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("AMTTRANSPORTTASK"), 2)%></font></td>
		      <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("amtsaldo"), 2)%></font></td>
			  <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("amtperunit"), 2)%></font></td>
		</tr>
		
<%
		  subopeningbalanceamt = subopeningbalanceamt + cdbl(rsCat("openingbalanceamt"))
		  subamtreceived = subamtreceived + cdbl(rsCat("amtreceived"))
		  subamtwriteoff = subamtwriteoff + cdbl(rsCat("amtwriteoff"))
		  subamtissued = subamtissued + cdbl(rsCat("amtissued"))
          subamtsaldo = subamtsaldo + cdbl(rsCat("amtsaldo"))
		  subamttransporttask= subamttransporttask + cdbl(rsCat("AMTTRANSPORTTASK"))
		  subamtreceivedtask= subamtreceivedtask + cdbl(rsCat("AMTRECEIVEDTASK"))
%>

<%End If%>
	  <%
          totalopeningbalanceqty = totalopeningbalanceqty + cdbl(rsCat("openingbalanceqty"))
          totalqtysaldo = totalqtysaldo + cdbl(rsCat("qtysaldo"))
		  
		  totalopeningbalanceamt = totalopeningbalanceamt + cdbl(rsCat("openingbalanceamt"))
		  totalamtreceived = totalamtreceived + cdbl(rsCat("amtreceived"))
		  totalamtwriteoff = totalamtwriteoff + cdbl(rsCat("amtwriteoff"))
		  totalamtissued = totalamtissued + cdbl(rsCat("amtissued"))
          totalamtsaldo = totalamtsaldo + cdbl(rsCat("amtsaldo"))
		  totalamttransporttask = totalamttransporttask + cdbl(rsCat("AMTTRANSPORTTASK"))
		  totalamtreceivedtask  = totalamtreceivedtask + cdbl(rsCat("AMTRECEIVEDTASK"))

	  rsCat.MoveNext

	  response.Flush()
	  no = no + 1
	  Prevcontroljobname= Currcontroljobname
	Wend
	rsCat.close()
	rsCat=null%>
		<tr bgcolor="#9CB770">
			<td align="left" colspan="12" style="padding-left:10px;"><font face="Verdana" size="1"><strong>Sub Total <%=PCase(Prevcontroljobname)%></strong></font></td>
			<td align="Right"><font face="Verdana" size="1"><strong><%=FormatNumber(subopeningbalanceamt, 2)%></strong></font></td>
			<td align="Right"><font face="Verdana" size="1"><strong><%=FormatNumber(subamtreceived, 2)%></strong></font></td>
			<td align="Right"><font face="Verdana" size="1"><strong><%=FormatNumber(subamtreceivedtask, 2)%></strong></font></td>
			<td align="Right"><font face="Verdana" size="1"><strong><%=FormatNumber(subamtwriteoff, 2)%></strong></font></td>
			<td align="Right"><font face="Verdana" size="1"><strong><%=FormatNumber(subamtissued, 2)%></strong></font></td>
			<td align="Right"><font face="Verdana" size="1"><strong><%=FormatNumber(subamttransporttask, 2)%></strong></font></td>
			<td align="Right"><font face="Verdana" size="1"><strong><%=FormatNumber(subamtsaldo, 2)%></strong></font></td>
			<td align="Right"><font face="Verdana" size="1"><strong>&nbsp;</strong></font></td>
		</tr>
		<tr style="background-color:#9CB770;">
			<td align="Left" colspan="12" style="padding-left:10px;"><font face="Verdana" size="1"><strong>Grand Total</strong></font></td>
		      <!--<td align="right"><font face="Verdana" size="1"><%=FormatNumber(totalopeningbalanceqty, 2)%></font></td>
		      <td align="right"><font face="Verdana" size="1"><%=FormatNumber(0, 2)%></font></td>
		      <td align="right"><font face="Verdana" size="1"><%=FormatNumber(0, 2)%></font></td>
		      <td align="right"><font face="Verdana" size="1"><%=FormatNumber(0, 2)%></font></td>
		      <td align="right"><font face="Verdana" size="1"><%=FormatNumber(totalqtysaldo, 2)%></font></td>-->
		      <td align="right"><font face="Verdana" size="1"><strong><%=FormatNumber(totalopeningbalanceamt, 2)%></strong></font></td>
		      <td align="right"><font face="Verdana" size="1"><strong><%=FormatNumber(totalamtreceived, 2)%></strong></font></td>
			  <td align="right"><font face="Verdana" size="1"><strong><%=FormatNumber(totalamtreceivedtask, 2)%></strong></font></td>
		      <td align="right"><font face="Verdana" size="1"><strong><%=FormatNumber(totalamtwriteoff, 2)%></strong></font></td>
		      <td align="right"><font face="Verdana" size="1"><strong><%=FormatNumber(totalamtissued, 2)%></strong></font></td>	
              <td align="right"><font face="Verdana" size="1"><strong><%=FormatNumber(totalamttransporttask, 2)%></strong></font></td>				  
			  <td align="right"><font face="Verdana" size="1"><strong><%=FormatNumber(totalamtsaldo, 2)%></strong></font></td>
		      <td align="right"><font face="Verdana" size="1"><strong>&nbsp;</strong></font></td>
		</tr>


	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
	</tbody>
</table>
<!--#include file="Bottom_Report.asp"-->
		</DIV>
	</BODY>
</HTML>
	