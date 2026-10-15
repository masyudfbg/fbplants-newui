<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>ePlaS Reporting - Sales Invoice DP</title>
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
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
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

v_transdate_from=trim(request("transdate_from"))
v_transdate_to=trim(request("transdate_to"))
v_txtSALESCONTRACTNO=trim(request("SALESCONTRACTNO"))
invoicedpno=trim(request("invoicedpno"))
' v_txtITEMCODE=trim(request("ITEMCODE"))
' v_txtRequestorID=trim(request("REQUESTORID"))
' v_txtRowStateId=trim(request("ROWSTATEID"))
businessunit= trim(request("businessunit"))
company     = trim(request("company"))

q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & businessunit & "' "
	' response.write q
set rsCat = connStr.Execute(q)
if not rsCat.eof then
   fcbaname = rsCat("fcName")
   CompanyName=rsCat("FCCOMPANYNAME")
   MgrName = rsCat("FCMANAGER")
   KTUName = rsCat("FCKTU")
   CropName=rsCat("FCCROP")
   FcbaType=rsCat("fcType")
end if


txtSQL = 	" SELECT fcba, invoicedpno," & _
			"        transdate," & _
			"        salescontractno," & _
			"        nvl(totalamount,0)totalamount," & _
			"        nvl(dp_beforevat,0)dp_beforevat," & _
			"        nvl(dp_vat,0)dp_vat," & _
			"        nvl(amount_dpvat,0)amount_dpvat," & _
			"        nvl(todate_dp_aftervat,0)todate_dp_aftervat," & _
			"        nvl(percentage_dp,0)percentage_dp," & _
			"        lineno," & _
			"        nvl(todate_percentage_dp,0)todate_percentage_dp," & _
			"        valid_until" & _
			"   FROM salesinvoicedp "  

if  len(trim(v_transdate_from)) > 0 then
    txtSQL = txtSQL & " WHERE transdate >=to_Date('" & v_transdate_from & "','MM/DD/YYYY')"
	txtSelect = txtSelect & " From " & FormatDateTime(v_transdate_from,1)	
end if
if  len(trim(v_transdate_to)) > 0 then
    txtSQL =txtSQL &  " and transdate <=to_Date('" & v_transdate_to & "','MM/DD/YYYY')"
	txtSelect = txtSelect & " To " & FormatDateTime(v_transdate_to,1)	
end if
if  len(trim(v_txtSALESCONTRACTNO)) > 0 then
    txtSQL = txtSQL & " and salescontractno='" & v_txtSALESCONTRACTNO & "'"
	txtSelect = txtSelect & " Sales Contract No. " & v_txtSALESCONTRACTNO	
end if
if  len(trim(invoicedpno)) > 0 then
    txtSQL = txtSQL & " and invoicedpno='" & invoicedpno & "'"
	txtSelect = txtSelect & " Sales Invoice DP No. " & invoicedpno	
end if
' if  len(trim(businessunit)) > 0 then
     txtSQL = txtSQL & " and fcba in (SELECT fccode from BusinessUnit where fccompanycode='" & company & "' and fccode like '%"& businessunit &"%')" 
' end if

	txtSQL = txtSQL & " Order by invoicedpno, transdate, SALESCONTRACTNO " 'ditambah ukie tgl 9 okt 2013 untuk Order by


if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
end if
' response.Write(txtsql)
response.Flush()
%>
<body> 
	<div id="wrapper">
	<!--#include file="print_excel.asp"-->
		<table border="0" width="100%" id="product-table" class="xthf-site">
			<thead>
				<tr>
					<th colspan="13" style=" background-color:#FFFFFF; color:#000000; border:none">
						<div id="header-wrapper">
							<div id="head_judul" style="float:left; text-align:left;">
								<h4>PT. <%=CompanyName%></h4>
								<h5><%=fcbaname%></h5>
							</div>
							<div id="subheader" class="centered" style="width:50%; margin:auto;">
								<h2>SALES INVOICE DP NO</h2>
							</div>
						</div>
					</th>
				</tr>
				<!--<div class="printhide">-->
				<tr>
					<th colspan="13" style=" background-color:#FFFFFF; color:#000000; border:none"><%=txtSelect%></td>
				</tr>
				<!--</div>-->
				<tr>
					<%	
					Set rsCat = Server.CreateObject("ADODB.Recordset")
					rsCat.CursorType = 3
					rsCat.CursorLocation = 3
					rsCat.Open txtSQL, connstr


					if Not rsCat.eof then
						nCountofField=rsCat.fields.count-1
					end if
					NO=1%>
				</tr>
				<tr>
					<th>Sales Invoice DP No</th>
					<th>Date</th>
					<th>Sales Contract No</th>
					<th>Total Amount</th>
					<th>Total DP Before VAT</th>
					<th>DP VAT (%)</th>
					<th>Amount DP VAT</th>
					<th>To Date DP After VAT</th>
					<th>Percentage DP (%)</th>
					<th>To Date Pecentage DP (%)</th>
					<th>Valid Until</th>
				</tr>
			</thead>
			<tbody>
				<%
					Dim v_Opening,v_Debit,v_Credit,v_Saldo,v_Quantity
					v_Quantity=0
					v_Amount = 0
					While Not rsCat.EOF
				%>
				<tr>
					<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1">
						<a href="javascript:launchwin('print_salesinvoicedp.Asp?invoicecode=<%=rscat("invoicedpno")%>&transdate=<%=rscat("transdate")%>&salescontractno=<%=rscat("salescontractno")%>&businessunit=<%=rscat("fcba")%>&Mode=Print','newwindow','width='+(980)+',height='+(650)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">
							<font face="Verdana" size="1"><%=rscat("invoicedpno")%></font>
						</a>
					</td>
					<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("transdate")%> </td>
					<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("salescontractno")%> </td>
					<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("totalamount")),2)%></td>
					<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("dp_beforevat")),2)%></td>
					<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("dp_vat")),2)%></td>
					<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("amount_dpvat")),2)%></td>
					<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("todate_dp_aftervat")),2)%></td>
					<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("percentage_dp")),2)%></td>
					<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("todate_percentage_dp")),2)%></td>
					<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("valid_until")%> </td>
				</tr>
				<%
					rsCat.MoveNext
					'response.Flush()
					no=no+1
					Wend
					rsCat.close()
					rsCat=null
				%>
				<input  type="hidden" value="<%=txtTable%>" name="txtTable">
				<input  type="hidden" value="<%=txtModule%>" name="txtModule">
			</tbody>
		</table>
	<!--#include file="Bottom_Report.asp"-->
	</DIV>
</body>	
</HTML>