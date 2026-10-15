<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>ePlaS Report - Purchase Contract</title>
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
d_span = 26
creationdate_from = request("creationdate_from")
creationdate_to = request("creationdate_to")
purchasecontractno = request("purchasecontractno")
suppliercode = request("suppliercode")
productcode = request("productcode")
deliverytype = request("deliverytype")

txtSelect = ""
businessunit = request("businessunit")

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


txtSQL = " SELECT ddata.*, cs.fcname customername, pr.fcname deliverytypename" & _
		 " FROM view_purchaseagreement ddata" & _
		 " LEFT JOIN supplier cs ON (ddata.fcba = cs.fcba AND ddata.suppliercode = cs.fccode)" & _
		 " LEFT JOIN parameter pr ON (ddata.fcba = pr.fcba AND ddata.deliverytype = pr.fccode AND pr.parcode = 'DELIVERY_TYPE')" & _
		 " WHERE ddata.creationdate BETWEEN TO_DATE('" & creationdate_from & "','mm/dd/yyyy') AND TO_DATE('" & creationdate_to & "','mm/dd/yyyy')" & _
		 "       AND ddata.fcba = '" & businessunit & "'"

if cstr(""&creationdate_from) = cstr(""&creationdate_to) then
	txtSelect = txtSelect & " Date : " & Day(creationdate_to) & " " & MonthName(Month(creationdate_to)) & " " & Year(creationdate_to)
else
	txtSelect = txtSelect & " From " & Day(creationdate_from) & " " & MonthName(Month(creationdate_from)) & " " & Year(creationdate_from)
	txtSelect = txtSelect & " To " & Day(creationdate_to) & " " & MonthName(Month(creationdate_to)) & " " & Year(creationdate_to)
end if

if cstr(""&purchasecontractno) <> "" then
    txtSQL = txtSQL & " AND ddata.purchasecontractno = '" & purchasecontractno & "'"
	txtSelect = txtSelect & " Contract No " & purchasecontractno
end if

if cstr(""&suppliercode) <> "" then
    txtSQL = txtSQL & " AND ddata.suppliercode = '" & suppliercode & "'"
	txtSelect = txtSelect & " Supplier " & suppliercode
end if

if cstr(""&productcode) <> "" then
    txtSQL = txtSQL & " AND ddata.productcode = '" & productcode & "'"
	txtSelect = txtSelect & " Product " & productcode
end if

if cstr(""&deliverytype) <> "" then
    txtSQL = txtSQL & " AND ddata.deliverytype = '" & deliverytype & "'"
	txtSelect = txtSelect & " Delivery Type " & deliverytype
end if

if len(trim(cstr(""&txtSelect))) > 0 then
   txtSelect = "" & txtSelect
else
   txtSelect = "Selected by All Data"
end if

arr_data_eof = true

' response.write txtSQL
Set rsCat = connStr.Execute(txtSQL)
if Not rsCat.eof then
	arr_data = rsCat.getrows()
	arr_data_eof = false
end if

%>
	<body>
		<div id="wrapper">
        <!--#include file="print_excel.asp"-->
        <table border="0" width="100%" id="product-table" class="xthf-site">
		<thead>
		<tr>
			<th colspan="<%=d_span%>" style=" background-color:#FFFFFF; color:#000000; border:none">
				<div id="header-wrapper">
					<div id="head_judul" style="float:left; text-align:left;">
						<h4>PT. <%=CompanyName%></h4>
						<h5><%=fcbaname%></h5>
					</div>
					<div id="subheader" class="centered" style="width:50%; margin:auto;">
						<h2>Purchase Contract</h2>
					</div>
				</div>
			</th>
		</tr>
		<tr>
			<th colspan="<%=d_span%>" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
		</tr>
		<tr>
			<th>Contract No</th>
			<th>Creation Date</th>
			<th>Supplier</th>
			<th>Sales Contract No</th>
			<th>Delivery Address</th>
			<th>Delivery Type</th>
			<th>Currency</th>
			<th>Payment Condition</th>
			<th>Delivery Condition</th>
			<th>Remarks</th>
			<th>Ppn</th>
			<th>Down Before Payment</th>
			<th>Down After Payment</th>
			<th>Down Payment PPn</th>
			<th>Amount Down Payment</th>
			<th>Product</th>
			<th>Dirt</th>
			<th>Moisture & Impurity (Max) %</th>
			<th>FFA</th>
			<th>Dobi</th>
			<th>Uom</th>
			<th>Start</th>
			<th>End</th>
			<th>Qty</th>
			<th>Price</th>
			<th>Amount</th>
			<th>Total Amount</th>
		</tr>
		</thead>
		<tbody>
<%
		if not arr_data_eof then
			gt_qty = 0
			gt_amount = 0
			gt_vat = 0
			gt_totalamount = 0
			for xyy=0 to ubound(arr_data,2)
				response.write "<tr>"
				response.write "	<td align='center'>"&arr_data(0,xyy)&"</td>"
				response.write "	<td align='center'>"&arr_data(1,xyy)&"</td>"
				response.write "	<td align='center'>"&arr_data(3,xyy)& "-" & arr_data(33,xyy) &"</td>"
				response.write "	<td align='center'>"&arr_data(32,xyy)&"</td>"
				response.write "	<td align='center'>"&arr_data(4,xyy)&"</td>"
				response.write "	<td align='center'>"&arr_data(5,xyy)&"</td>"
				response.write "	<td align='center'>"&arr_data(6,xyy)&"</td>"
				response.write "	<td align='center'>"&arr_data(7,xyy)&"</td>"
				response.write "	<td align='center'>"&arr_data(8,xyy)&"</td>"
				response.write "	<td align='left'>"&arr_data(9,xyy)&"</td>"
				response.write "	<td align='right'>"&FormatNumber(cdbl(to_nol(arr_data(23,xyy))),2)&"</td>"
				response.write "	<td align='right'>"&FormatNumber(cdbl(to_nol(arr_data(10,xyy))),2)&"</td>"
				response.write "	<td align='right'>"&FormatNumber(cdbl(to_nol(arr_data(11,xyy))),2)&"</td>"
				response.write "	<td align='right'>"&FormatNumber(cdbl(to_nol(arr_data(12,xyy))),2)&"</td>"
				response.write "	<td align='right'>"&FormatNumber(cdbl(to_nol(arr_data(14,xyy))),2)&"</td>"
				response.write "	<td align='center'>"&arr_data(16,xyy)&"</td>"
				response.write "	<td align='center'>"&arr_data(28,xyy)&"</td>"
				response.write "	<td align='center'>"&arr_data(29,xyy)&"</td>"
				response.write "	<td align='center'>"&arr_data(30,xyy)&"</td>"
				response.write "	<td align='center'>"&arr_data(31,xyy)&"</td>"
				response.write "	<td align='right'>"&arr_data(17,xyy)&"</td>"
				response.write "	<td align='right'>"&arr_data(18,xyy)&"</td>"
				response.write "	<td align='right'>"&arr_data(19,xyy)&"</td>"
				response.write "	<td align='right'>"&FormatNumber(cdbl(to_nol(arr_data(20,xyy))),2)&"</td>"
				response.write "	<td align='right'>"&FormatNumber(cdbl(to_nol(arr_data(21,xyy))),2)&"</td>"
				response.write "	<td align='right'>"&FormatNumber(cdbl(to_nol(arr_data(22,xyy))),2)&"</td>"
				response.write "	<td align='right'>"&FormatNumber(cdbl(to_nol(arr_data(15,xyy))),2)&"</td>"
				response.write "</tr>"
				
				gt_dp_aftervat = gt_dp_aftervat + cdbl(to_nol(arr_data(11,xyy)))
				gt_dp_beforevat = gt_dp_beforevat + cdbl(to_nol(arr_data(10,xyy)))
				gt_amount_vat = gt_amount_vat + cdbl(to_nol(arr_data(14,xyy)))
				gt_qty = gt_qty + cdbl(to_nol(arr_data(20,xyy)))
				gt_amount = gt_amount + cdbl(to_nol(arr_data(22,xyy)))
				gt_vat = gt_vat + cdbl(to_nol(arr_data(23,xyy)))
				gt_totalamount = gt_totalamount + cdbl(to_nol(arr_data(15,xyy)))
			next
			response.write "<tr bgcolor='#F4F4A0'>"
				response.write "	<td align='center' colspan='11'><b>Total</b></td>"
				response.write "	<td align='right'><b>"&FormatNumber(gt_dp_beforevat,2)&"</b></td>"
				response.write "	<td align='right'><b>"&FormatNumber(gt_dp_aftervat,2)&"</b></td>"
				response.write "	<td align='right' colspan='10'><b></b></td>"
				response.write "	<td align='right'><b>"&FormatNumber(gt_qty,2)&"</b></td>"
				response.write "	<td align='right'></td>"
				response.write "	<td align='right'><b>"&FormatNumber(gt_amount,2)&"</b></td>"
				response.write "	<td align='right'><b>"&FormatNumber(gt_totalamount,2)&"</b></td>"
			response.write "</tr>"
		else
			response.write "<tr><td align='left' colspan='"&d_span&"'>No Data Available</td>"
		end if
		
%>
		</tbody>
		</table>
		</div>
		<input  type="hidden" value="<%=txtTable%>" name="txtTable">
		<input  type="hidden" value="<%=txtModule%>" name="txtModule">
		</form>
	</body>
</html>