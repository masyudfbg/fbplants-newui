<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
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
<title>ePlaS Reporting - Contractor Invoice </title>
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

 '---- Data Details

 fcba =Request("fcba")
 agreementcode =Request("agreementcode")

		q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & fcba & "' "
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

txtSQL = 	"   SELECT cih.invoicecode," & _
			"          cih.invoice_by," & _
			"          cih.transdate," & _
			"          cih.agreementcode," & _
			"          cih.contractorcode," & _
			"          contractor_api.Get_ContractorName(cih.contractorcode,cih.fcba) contractorname," & _ 
			"          cih.isretention_invoice," & _
			"          cih.retention," & _
			"          cih.totalamount_afterretention," & _
			"          cih.ppn," & _
			"          cih.amountppn," & _
			"          cih.pph," & _
			"          cih.amountpph," & _
			"          cih.grandtotal_aftertax," & _
			"          cih.description," & _
			"          cih.rowstate," & _
			"          cih.fcba," & _
			"          cid.jobcode," & _
			"           (cid.rate) rate," & _
			"          SUM (cid.totalfinal) totalfinal," & _
			"          SUM (cid.amount) amount," & _
			"          SUM (cid.amountretention) amountretention," & _
			"          SUM (cid.amountafter_retention) amountafter_retention," & _
			"          cid.remarks" & _
			"     FROM contractorinvoicehdr cih" & _
			"          INNER JOIN contractorinvoicedtl cid" & _
			"             ON (    cih.invoicecode = cid.invoicecode" & _
			"                 AND cih.agreementcode = cid.agreementcode" & _
			"                 AND cih.fcba = cid.fcba)" & _
			" WHERE cih.fcba='"& fcba &"' and cih.agreementcode = '"& agreementcode &"'" 
			

txtsql = txtSQL &	"  GROUP BY cih.invoicecode," & _
					"          cih.invoice_by," & _
					"          cih.transdate," & _
					"          cih.agreementcode," & _
					"          cih.contractorcode," & _
					"          cih.isretention_invoice," & _
					"          cih.retention," & _
					"          cih.totalamount_afterretention," & _
					"          cih.ppn," & _
					"          cih.amountppn," & _
					"          cih.pph," & _
					"          cih.amountpph," & _
					"          cih.grandtotal_aftertax," & _
					"          cih.description," & _
					"          cih.rowstate," & _
					"          cih.fcba," & _
					"          cid.jobcode," & _
					"          cid.rate," & _
					"          cid.remarks" 

txtsql = txtSQL & " ORDER BY invoicecode, transdate, fcba"


' response.Write(txtsql)
' response.Flush()
Set rsCat = Server.CreateObject("ADODB.Recordset")
	rsCat.CursorType = 3
	rsCat.CursorLocation = 3
	rsCat.Open txtSQL, connstr
%>
 <body>
		<div id="wrapper">
		<!--#include file="print_excel.asp"-->
 <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="20" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=CompanyName%></h4>
					<h5><%=fcbaname%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>CONTRACTOR INVOICE </h2>
				</div>
			</div>
		 </th>
			</tr>
						<!--<div class="printhide">-->
				<tr>
					<td colspan="20" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></td>
				</tr>
				<tr>
					<th>Invoice No</th>
					<th>Invoice By</th>
					<th>Agreement Code</th>
					<th>Date</th>
					<th>Contractor Code</th>
					<th>Contractor Name</th>
					<th>Is Retention Invoice</th>
					<th>Retention (%)</th>
					<th>PPn (%)</th>
					<th>Amount PPn</th>
					<th>PPh 23</th>
					<th>Amount PPh 23</th>
					<th>Rowstate</th>
					<th>Jobcode</th>
					<th>Total Final</th>
					<th>Rate</th>
					<th>Amount</th>
					<th>Amount Retention</th>
					<th>Amount After Retention</th>
					<th>Remarks</th>
				</tr>
                </thead>
               <tbody>
	<%
	if Not rscat.eof then
		v_totalamount_afterretention = 0
		v_amountppn = 0
		v_amountpph = 0
		v_grandtotal_aftertax = 0
		v_totalfinal = 0
		v_rate = 0
		v_amount = 0
		v_amountretention = 0
		v_amountafter_retention = 0
	While Not rscat.EOF


	%>
		<tr>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("invoicecode")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("invoice_by")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("agreementcode")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("transdate")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("contractorcode")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("contractorname")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("isretention_invoice")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("retention")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("ppn")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rscat("amountppn"))),0)%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("pph")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rscat("amountpph"))),0)%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("rowstate")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("jobcode")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rscat("totalfinal"))),2)%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rscat("rate"))),0)%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rscat("amount"))),0)%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rscat("amountretention"))),0)%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rscat("amountafter_retention"))),0)%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("remarks")%></td>
		</tr>
	  <%
	   ' v_totalamount_afterretention =v_totalamount_afterretention+trim(cdbl(rsCat("totalamount_afterretention")))
	   v_amountppn =v_amountppn+Trim(cdbl(rsCat("amountppn")))
	   v_amountpph =v_amountpph+Trim(cdbl(rsCat("amountpph")))
	   ' v_grandtotal_aftertax =v_grandtotal_aftertax+trim(cdbl(rsCat("grandtotal_aftertax")))
	   v_totalfinal =v_totalfinal+Trim(cdbl(rsCat("totalfinal")))
	   v_rate =v_rate+Trim(cdbl(rsCat("rate")))
	   v_amount =v_amount+trim(cdbl(rsCat("amount")))
	   v_amountretention =v_amountretention+Trim(cdbl(rsCat("amountretention")))
	   v_amountafter_retention =v_amountafter_retention+Trim(cdbl(rsCat("amountafter_retention")))
	  rsCat.MoveNext
	Wend
	Else
	response.write "<tr><td colspan='19'>No Data Available</td></tr>"
	End If
	rsCat.close()
	rsCat=null%>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
	<tr bgcolor="#F4F4A0">
		<td align="Center" style="<%=cmdstyleLeft%>" colspan="9"> <font face="Verdana" size="1">T o t a l</td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(v_amountppn)),0)%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(v_amountpph)),0)%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(v_totalfinal)),0)%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(v_rate)),0)%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(v_amount)),0)%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(v_amountretention)),0)%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(v_amountafter_retention)),0)%></td>
		<td align="Right" style="<%=cmdstyle%>"></td>
	</tr>
	</tbody>
</table>
<!--#include file="Bottom_Report.asp"-->	