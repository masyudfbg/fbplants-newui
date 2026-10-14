<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>ePlaS Reporting - Stock Write Off</title>
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
<body>
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

v_txtWODATE_From=Trim(REQUEST("WODATE_From"))
v_txtWODATE_To=Trim(REQUEST("WODATE_To"))
v_txtWOCODE=Trim(REQUEST("WriteOffCode"))
v_txtSTOCKCODE=Trim(REQUEST("STOCKCODE"))
v_txtLOCATIONTYPE=Trim(REQUEST("LOCATIONTYPE"))
v_txtLOCATIONCODE=Trim(REQUEST("LOCATIONCODE"))

txtSQL = " select wod.writeoffcode" & _
         "        , wo.storecode" & _
		 "        , to_char(wo.wodate, 'mm/dd/yyyy') wodate" & _
		 "        , wod.linenokey" & _
		 "        , wod.linenohist" & _
		 "        , wod.stockcode" & _
		 "        , stock_api.get_stockName(wod.stockcode, wod.fcba) stockdesc" & _
		 "        , wod.jobcode" & _
		 "        , job_api.get_jobcode_name(wod.jobcode, wod.fcba) jobdesc" & _
		 "        , wod.locationtype" & _
         "        , wod.locationcode" & _
		 "        , wod.reference" & _
         "        , wod.remarks" & _
         "        , nvl(sum(wod.quantity), 0) quantity" & _
         "        , nvl(sum(wod.price), 0) price" & _
         "        , nvl(sum(wod.amount), 0) amount" & _
		 " from writeoff wo" & _
         "      , writeoffdetails wod" & _
		 " where wo.writeoffcode = wod.writeoffcode" & _
		 "       and wo.fcba = wod.fcba" & _
		 "       and wod.fcba = '" & session("fcBA") & "'"

if  len(trim(v_txtWODATE_From)) > 0 then
    txtSQL = txtSQL & " and wo.wodate >= to_date('" & v_txtWODATE_From & "', 'MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(v_txtWODATE_From,1)
end if
if  len(trim(v_txtWODATE_To)) > 0 then
    txtSQL = txtSQL & " and wo.wodate <= to_date('" & v_txtWODATE_To & "', 'MM/DD/YYYY')"
	txtSelect = txtSelect & "  to " & FormatDateTime(v_txtWODATE_To,1)
end if
if  len(trim(v_txtWOCODE)) > 0 then
    txtSQL = txtSQL & " and wod.writeoffcode = '" & v_txtWOCODE & "'"
	txtSelect = txtSelect & " Writeoff Code " & v_txtWOCODE
end if
if  len(trim(v_txtSTOCKCODE)) > 0 then
    txtSQL = txtSQL & " and wod.stockcode = '" & v_txtSTOCKCODE & "'"
	txtSelect = txtSelect & " Stock Code " & v_txtSTOCKCODE
end if 
if  len(trim(v_txtLOCATIONTYPE)) > 0 then
    txtSQL = txtSQL & " and wod.locationtype = '" & v_txtLOCATIONTYPE & "'"
	txtSelect = txtSelect & " Target Type " & v_txtLOCATIONTYPE
end if 
if  len(trim(v_txtLOCATIONCODE)) > 0 then
    txtSQL = txtSQL & " and wod.locationcode = '" & v_txtLOCATIONCODE & "'"
	txtSelect = txtSelect & " Target Code " & v_txtLOCATIONCODE
end if 

txtsql = txtsql & " group by wod.writeoffcode" & _
                  "          , wo.storecode" & _
                  "          , to_char(wo.wodate, 'mm/dd/yyyy')" & _
                  "          , wod.linenokey" & _
                  "          , wod.linenohist" & _
                  "          , wod.stockcode" & _
                  "          , stock_api.get_stockName(wod.stockcode, wod.fcba)" & _
                  "          , wod.jobcode" & _
                  "          , job_api.get_jobcode_name(wod.jobcode, wod.fcba)" & _
                  "          , wod.locationtype" & _
                  "          , wod.locationcode" & _
                  "          , wod.reference" & _
                  "          , wod.remarks"
txtsql = txtsql & " order by wod.writeoffcode" & _
                  "          , to_char(wo.wodate, 'mm/dd/yyyy')" & _
                  "          , wod.linenokey"
if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
end if
'response.Write txtsql
'response.Flush()
%>
	<body>
		<div id="wrapper">
        <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="13" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=Session("COMPANY")%></h4>
					<h5><%=Session("BA")%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>STOCK WRITE OFF</h2>
				</div>
			</div>
				</th>
			</tr>
						<!--<div class="printhide">-->
						<tr>
							<td colspan="13" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></td>
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
							<tr>
							    
							   <th>Write Off Code</th>
							   <th>Write Off Date</th>
							   <th>Stock Code</th>
							   <th>Stock Description</th>
							   <th>Job Code</th>
							   <th>Job Description</th>
							   <th>Target Type</th>
							   <th>Target Code</th>
							   <th>Line No History</th>
							   <th>Reference</th>
							   <th>Remarks</th>
							   <th>Quantity</th>
							   <th>Amount</th>
							</tr>
                             </thead>
                            <tbody>
								<%
								Dim v_Quantity,v_price,v_amount
								v_Quantity=0
								v_price=0
								v_amount=0
								While Not rsCat.EOF%>
							<tr>
								   
								  <td align="left" style="<%=cmdstyle%>" title="PO.ITEMCODE" valign="top"><font face="Verdana" size="1"><%=rsCat("WriteOffCODE")%> </td>
								  <td align="left" style="<%=cmdstyle%>" title="PO.ITEMCODE" valign="top"><font face="Verdana" size="1"><%=rsCat("wodate")%> </td>
								  <td align="left" style="<%=cmdstyle%>" title="PO.OTHERITEMDESC" valign="top"><font face="Verdana" size="1"><%=rsCat("STOCKCODE")%> </td>
								  <td align="left" style="<%=cmdstyle%>" title="PO.OTHERITEMDESC" valign="top"><font face="Verdana" size="1"><%=rsCat("StockDesc")%> </td>
								  <td align="left" style="<%=cmdstyle%>" title="s.fcname" valign="top"><font face="Verdana" size="1"><%=rsCat("JOBCODE")%></td>
								  <td align="left" style="<%=cmdstyle%>" title="PO.OTHERITEMDESC" valign="top"><font face="Verdana" size="1"><%=rsCat("JobDesc")%> </td>
								  <td align="left" style="<%=cmdstyle%>" title="PO.OTHERITEMDESC" valign="top"><font face="Verdana" size="1"><%=rsCat("LOCATIONTYPE")%> </td>
								  <td align="left" style="<%=cmdstyle%>" title="PO.OTHERITEMDESC" valign="top"><font face="Verdana" size="1"><%=rsCat("LOCATIONCODE")%> </td>
								  <td align="Right" style="<%=cmdstyle%>" title="PO.OTHERITEMDESC" valign="top"><font face="Verdana" size="1"><%=rsCat("linenohist")%> </td>
								  <td align="left" style="<%=cmdstyle%>" title="PO.OTHERITEMDESC" valign="top"><font face="Verdana" size="1"><%=rsCat("REFERENCE")%> </td>
								  <td align="left" style="<%=cmdstyle%>" title="PO.OTHERITEMDESC" valign="top"><font face="Verdana" size="1"><%=rsCat("REMARKS")%> </td>
								  <td align="Right" style="<%=cmdstyle%>" title="po.Quantity" valign="top"><font face="Verdana" size="1"><%=FormatNumber(rsCat("Quantity"), 2)%> </td>
								  <td align="Right" style="<%=cmdstyle%>" title="po.Amount" valign="top"><font face="Verdana" size="1"><%=FormatNumber(rsCat("AMOUNT"), 2)%> </td>
							</tr>
							  <%
							  v_Quantity = v_Quantity + trim(cdbl(rsCat("quantity")))
							  v_price = v_price + Trim(cdbl(rsCat("price")))
							  v_amount = v_amount + Trim(cdbl(rsCat("amount")))
							  'response.write cdbl(trim(rsCat("MTHOPENBALANCE")))+1
							  rsCat.MoveNext
							  'response.Flush()
							  no=no+1
							Wend
							rsCat.close()
							rsCat=null%>
							<input  type="hidden" value="<%=txtTable%>" name="txtTable">
							<input  type="hidden" value="<%=txtModule%>" name="txtModule">
							</form>
							<tr bgcolor="#F4F4A0">
							  <td colspan="11" style="<%=cmdstyleLeft%>" align="center"><font face="Verdana" size="1"><strong>T o t a l</strong></td>
							  <td align="right"style="<%=cmdstyle%>" ><font face="Verdana" size="1"><strong><%=FormatNumber(v_Quantity,2)%></strong></td>
							  <td align="right"style="<%=cmdstyle%>" ><font face="Verdana" size="1"><strong><%=FormatNumber(v_amount,2)%></strong></td>
							</tr>
                            </tbody>
						</table>
						<!--#include file="Bottom_Report.asp"-->
		</DIV>
	</BODY>
</HTML>	