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
<title>ePlaS Reporting - Service Purchase Order </title>
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

 ' End Added by NAF
 ITEMCODE=Request("ITEMCODE")
 QUANTITY=Request("QUANTITY")
 AMOUNT =Request("AMOUNT")
 OTHERITEMDESC=request("OTHERITEMDESC")
 RowStateId = Request("RowStateId") 
 
 PRCODE = Request("PRCODE") 
 PRLINEID = Request("PRLINEID")
 use_ppn = Request("use_ppn")
 fccompanycode = Request("fccompanycode")
 businessunit = Request("businessunit")
 purchaseby = Request("purchaseby")
 suppliername = Request("suppliername")
 supplieraddress = Request("supplieraddress")
 termofpayment = Request("termofpayment")
 SHIPPINGPOINT = Request("SHIPPINGPOINT")
 dp = Request("dp")
 ppn = Request("ppn")
 deliverydestination = Request("deliverydestination")
 StockName = Request("StockName")
 requestedqty = Request("requestedqty")
 unitprice = Request("unitprice")
 disc = Request("disc")
 othercharge = Request("othercharge")
 totalamount = Request("totalamount")
 totalamountppn = Request("totalamountppn")
 totalcharge = Request("totalcharge")
 grandtotal = Request("grandtotal")
 amountcharge = Request("amountcharge")
 totalppn = Request("totalppn")
 approvedqty = request ("approvedqty")
 podate_from = Request("podate_from")
 podate_to = Request ("podate_to")
 spono = Request ("spono")

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

txtSQL = 	"	SELECT sph.company," & _
			"       sph.fcba," & _
			"       sph.suppliercode," & _
			"       sph.purchaseby," & _
			"       sph.termofpayment," & _
			"       sph.pocode," & _
			"       sph.podate," & _
			"       sph.rowstateid," & _
			"       spd.stockname," & _
			"       spd.quantity," & _
			"       spd.unitprice," & _
			" 		spd.disc, " & _
		    "	    spd.amountcharge, " & _
       		"		spd.othercharge, " & _
			"       sph.grandtotal" & _
			"  FROM servicepurchaseorder sph" & _
			"       INNER JOIN servicepurchaseorderdetails spd" & _
			"          ON (sph.pocode = spd.pocode AND sph.fcba = spd.fcba)" 

if  len(trim(fccompanycode)) > 0 then
    txtSQL =txtSQL &  "  where sph.company = '"& fccompanycode &"'"
end if
if  len(trim(businessunit)) > 0 then
    txtSQL =txtSQL &  "  and sph.fcba = '"& businessunit &"'"
end if
if  len(trim(purchaseby)) > 0 then
    txtSQL =txtSQL &  " and sph.purchaseby = '"& purchaseby &"'"
end if
if  len(trim(spono)) > 0 then
    txtSQL =txtSQL &  " and sph.POCODE = '"& spono &"'"
end if
if  len(trim(suppliercode)) > 0 then
    txtSQL =txtSQL &  " and sph.suppliercode = '"& suppliercode &"'"
end if
if  len(trim(termofpayment)) > 0 then
    txtSQL =txtSQL &  " and sph.termofpayment = '"& termofpayment &"'"
end if
if  len(trim(podate_from)) > 0 then
    txtSQL = txtSQL & " and podate >=to_Date('" & podate_from & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(podate_from,1)
end if
if  len(trim(podate_to)) > 0 then
    txtSQL =txtSQL &  " and podate <=to_Date('" & podate_to & "','MM/DD/YYYY')"
	txtSelect = txtSelect & "  to " & FormatDateTime(podate_to,1)
end if

txtsql=(txtsql) 
if len(trim(txtSelect)) > 0 then
   txtSelect = "" & txtSelect
else
   txtSelect = "Selected by All Data"
end if
'& " group by WIDATE, WIDIVISION"

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
         <th colspan="14" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=CompanyName%></h4>
					<h5><%=businessunit%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>SERVICE PURCHASE ORDER </h2>
				</div>
			</div>
		 </th>
			</tr>
						<!--<div class="printhide">-->
				<tr>
					<td colspan="14" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></td>
				</tr>
				<tr>
					<th>Company</th>
					<th>Business Unit</th>
					<th>Date</th>
					<th>SPO</th>
					<th>Status</th>
					<th>Purchase By</th>
					<th>Vendor</th>
					<th>Term Of Payment</th>
					<th>Item Name</th>
					<th>Quantity</th>
					<th>Disc (%)</th>
					<th>Unit Price</th>
					<th>Other Charge</th>
					<th>Amount</th>
				</tr>
                </thead>
               <tbody>
	<%
	if Not rscat.eof then
		v_unitprice = 0
		v_amountcharge = 0
		v_othercharge = 0
	While Not rscat.EOF


	%>
		<tr>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("company")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fcba")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("podate")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("pocode")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("RowStateId")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("purchaseby")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("suppliercode")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("termofpayment")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("stockname")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("quantity")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("disc")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(rscat("unitprice"),2)%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(rscat("othercharge"),2)%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(rscat("amountcharge"),2)%></td>
		</tr>
	  <%
	   v_unitprice =v_unitprice+trim(cdbl(rsCat("unitprice")))
	   v_amountcharge =v_amountcharge+Trim(cdbl(rsCat("amountcharge")))
	   v_othercharge =v_othercharge+Trim(cdbl(rsCat("othercharge")))
	  rsCat.MoveNext
	Wend
	Else
	response.write "<tr><td colspan='14'>No Data Available</td></tr>"
	End If
	rsCat.close()
	rsCat=null%>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
	<tr bgcolor="#F4F4A0">
		<td align="Center" style="<%=cmdstyleLeft%>" colspan="11"> <font face="Verdana" size="1">T o t a l</td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_unitprice,2)%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_othercharge,2)%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_amountcharge,2)%></td>
	</tr>
	</tbody>
</table>
   <table width="90%" class="data">
   	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr>
 		<td width="50%" align="center"><font face="Verdana" size="1">Approved by,</font></td>
		<td width="50%" align="center"><font face="Verdana" size="1">Prepared by,</font></td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=MgrName%></u></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=KTUName%></u></strong></td>
	   </tr>
	   	  <tr> 
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>MANAGER</i></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>KTU</i></strong></td>
	   </tr>
       </tbody>
   </table>
   </body>
   </html>