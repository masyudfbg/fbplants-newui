<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%validatesession%>
<html>
	<head>
		<title>Print PO</title>
		<link rel="stylesheet" href="/plugins/paper-css/paper.min.css">
		<style>
			@page { size: A4 }
			table {
				border-collapse: collapse;
				width:100%;
				font: 11px arial, sans-serif;
			}
			
			th, td {
				padding: 5px 5px;
			}
			td.sideonly {
				border-top: none;
				border-bottom: none;
			}
			td.noborder {border:none;}
			article{font: 12px arial, sans-serif;}
			#wrapper {
				float: right;
				margin-bottom:10px;
			}
			
			input[type="button"]{
				background: #fff;
				-webkit-border-radius: 4;
				-moz-border-radius: 4;
				border-radius: 4px;
				font-family: Arial;
				color: #000;
				font-size: 12px;
				padding: 10px 20px 10px 20px;
				text-decoration: none;
				cursor:pointer;
				display:block;
				border-style: solid;
				border-width: 1px;
			}
			input.tableToExcel[type="button"] {
				background: #5bb75b;
				background-image: -webkit-linear-gradient(top, #5bb75b, #5bb75b);
				background-image: -moz-linear-gradient(top, #5bb75b, #5bb75b);
				background-image: -ms-linear-gradient(top, #5bb75b, #5bb75b);
				background-image: -o-linear-gradient(top, #5bb75b, #5bb75b);
				background-image: linear-gradient(to bottom, #5bb75b, #5bb75b);
				-webkit-border-radius: 4;
				-moz-border-radius: 4;
				border-radius: 4px;
				font-family: Arial;
				color: #ffffff;
				font-size: 12px;
				padding: 7px 10px;
				text-decoration: none;
			}			
			input.tableToExcel[type="button"]:hover {
				background: #4da34d;
				background-image: -webkit-linear-gradient(top, #4da34d, #4da34d);
				background-image: -moz-linear-gradient(top, #4da34d, #4da34d);
				background-image: -ms-linear-gradient(top, #4da34d, #4da34d);
				background-image: -o-linear-gradient(top, #4da34d, #4da34d);
				background-image: linear-gradient(to bottom, #4da34d, #4da34d);
				text-decoration: none;
			}
			.clear {clear:both;}
			@media print {
				#wrapper,.clear{display:none;}
			}
			input[type="text"].isian {border:0px;box-shadow:none; padding:0px;font-size: 10px;color:#000;height: 50px;width:100%; text-align:center; font-weight: bold; font-size:12pt;  border: 0;
									  outline: 0;
									  background: transparent;
									  border-bottom: 1px solid black;}
		</style>
	</head>
	<body class="A4">
		<section class="sheet padding-10mm">
		<div id="wrapper"><!--#include file="print_excel.asp"--></div>
		<div class="clear"></div>
		<article id="product-table">
<%
 fcba = Request("fcba")
 pocode = request("pocode")
 deliverydestination = request("deliverydestination")
 purchaseby = request("purchaseby")
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

q = " SELECT po.fcba," & _
	"       po.pocode," & _
	"       po.podate," & _
	"       TO_CHAR(po.podate, 'DD') podate_dd," & _
	"       TO_CHAR(po.podate, 'MM') podate_mm," & _
	"       TO_CHAR(po.podate, 'YYYY') podate_YY," & _
	"       po.suppliercode spcode," & _
	"       sp.fcname spname," & _
	"		SUPPLIER_API.GET_BANKACCOUNT(po.suppliercode, po.fcba) bankaccount," & _
	"		SUPPLIER_API.GET_BANKNAME(po.suppliercode, po.fcba) bankname," & _
	"		SUPPLIER_API.GET_BANKOWNERSHIP(po.suppliercode, po.fcba) bankownership," & _
	"		USER_TAB_API.get_fullname('DEPT_HEAD_PURCH') divhead_ss," & _
	"       sp.address spaddress," & _
	"       sp.city spcity," & _
	"       sp.state spstate," & _
	"       sp.phone," & _
	"       sp.fax," & _
	"       po.remarks," & _
	"       po.rowstateid," & _
	"       po.company," & _
	"       pod.prcode," & _
	"       po.ppn," & _
	"       po.totalamountppn," & _
	"       po.supplieraddress," & _
	"       po.purchaseby," & _
	"       po.deliverydestination," & _
	"       po.dp," & _
	"       po.totalcharge," & _
	"       po.grandtotal," & _
	"       po.totalamount," & _
	"       po.totalppn," & _
	"		bu2.address addresspurchaseby," & _
	"       bu2.fcname fcnamepurchaseby," & _
	"       bu2.fcfax faxpurchaseby," & _
	"       bu2.fctelephone telephonepurchaseby," & _
	"       bu3.fcfax faxdestination," & _
	"       bu3.fctelephone telephonedestination," & _
	"       bu3.address addressdestination," & _
	"       bu3.fcname fcnamedestination" & _
	"  FROM purchaseorderdetails pod" & _
	"		INNER JOIN purchaseorder po ON (pod.pocode = po.pocode AND pod.fcba = po.fcba)" & _
	"       LEFT JOIN supplier sp" & _
	"          ON (po.suppliercode = sp.fccode AND po.fcba = sp.fcba)" & _
	"       LEFT JOIN (SELECT *" & _
	"                    FROM businessunit" & _
	"                   WHERE fccode = '"& purchaseby &"') bu2" & _
	"          ON (bu2.fccode = po.purchaseby)" & _
	"       LEFT JOIN (SELECT *" & _
	"                    FROM businessunit" & _
	"                   WHERE fccode = '"& deliverydestination &"') bu3" & _
	"          ON (bu3.fccode = po.deliverydestination)" & _
	" WHERE po.fcba = '"& fcba &"' AND po.pocode = '"& pocode &"'" 

' response.Write q	
Set oRecordset = Server.CreateObject("ADODB.Recordset")
oRecordset.Open q, connstr
if not oRecordset.eof then
	addresspurchaseby = oRecordset("addresspurchaseby")
	fcnamepurchaseby = oRecordset("fcnamepurchaseby")
	faxpurchaseby = oRecordset("faxpurchaseby")
	telephonepurchaseby = oRecordset("telephonepurchaseby")
	faxdestination = oRecordset("faxdestination")
	telephonedestination = oRecordset("telephonedestination")
	addressdestination = oRecordset("addressdestination")
	fcnamedestination = oRecordset("fcnamedestination")
	podate = oRecordset("podate")
	fax = oRecordset("fax")
	phone = oRecordset("phone")
	divhead_ss = oRecordset("divhead_ss")
	podate_dd = oRecordset("podate_dd")
	podate_mm = oRecordset("podate_mm")
	podate_yy = oRecordset("podate_yy")
	spcode = oRecordset("spcode")
	spname = oRecordset("spname")
	bankaccount = oRecordset("bankaccount")
	bankname = oRecordset("bankname")
	bankownership = oRecordset("bankownership")
	spaddress = oRecordset("spaddress")
	spcity = oRecordset("spcity")
	spstate = oRecordset("spstate")
	remarks = oRecordset("remarks")
	ppn = oRecordset("ppn")
	company = oRecordset("company")
	totalamountppn = oRecordset("totalamountppn")
	supplieraddress = oRecordset("supplieraddress")
	purchaseby = oRecordset("purchaseby")
	deliverydestination = oRecordset("deliverydestination")
	dp = oRecordset("dp")
	totalcharge = oRecordset("totalcharge")
	grandtotal = oRecordset("grandtotal")
	totalamount = oRecordset("totalamount")
	totalppn = oRecordset("totalppn")
	prcode = oRecordset("prcode")
End if
oRecordset.close
set oRecordset = Nothing

if "" & spcity <> "" and "" & spstate <> "" then add_minus = " - " end if

Set oRecordset = Server.CreateObject("ADODB.Recordset")
oRecordset.Open q, connstr
IF NOT oRecordset.EOF THEN
response.write "<table style='border: px solid #111;'>"
response.write "	<thead>"
response.write "	</thead>"
response.write "	<tbody>"
' response.write "		<tr>"
' response.write "			<td width='56.6666666667%' colspan='7' class='noborder'><p id='logo'><img src='css/login/images/client.png' height='80px'></p></td>"
' response.write "			<td width='22%' colspan='2' align='center'></td>"
' response.write "		</tr>"
response.write "		<tr>"
response.write "			<td width='56.6666666667%' colspan='7' class='noborder'><img src='css/login/images/client.png' height='80px'><p style='font-size:16pt;'><b>PT. "&CompanyName&"</b></p></td>"
response.write "			<td width='22%' colspan='2' align='left'><p style='font-size:18pt;'><b>PURCHASE ORDER</b></p></td>"
response.write "		</tr>"
response.write "		<tr>"
response.write "			<td colspan='7' class='noborder' style='font-size:10pt;'></td>"
response.write "			<td style='width:7%; font-size:10pt'>No : </td>"
response.write "			<td style='font-size:10pt;'>"&pocode&"</td>"
response.write "		</tr>"
response.write "		<tr>"
response.write "			<td colspan='7' class='noborder' style='font-size:10pt;'><p>"&addresspurchaseby&"</p><p>Phone. "&telephonepurchaseby&"</br>Fax. "& faxpurchaseby &"</p></td>"
response.write "			<td style='width:7%; font-size:10pt'></td>"
response.write "		</tr>"
response.write "		<tr>"
response.write "			<td style='width:7%; font-size:10pt' colspan='7'></td>" 
response.write "		</tr>"
response.write "		<tr>"
response.write "			<td class='noborder' style='border-top: 0px solid #111;width:7%; font-size:10pt'>VENDOR</td>"
response.write "			<td class='noborder' style='border-top: 0px solid #111; font-size:10pt' colspan='6'>: <b>"&spname&"</b></td>"
response.write "			<td style='font-size:10pt;'>PO DATE</td>"
response.write "			<td style='font-size:10pt;'>"&podate_dd&" "&bulan(podate_mm)&" "&podate_yy&"</td>"
response.write "		</tr>"
response.write "		<tr>"
response.write "			<td class='noborder' style='font-size:10pt;'>ADDRESS</td>"
response.write "			<td class='noborder' colspan='6' style='font-size:10pt;'>: "&spaddress&"</td>"
response.write "			<td style='font-size:10pt;'>No. PR</td>"
response.write "			<td style='font-size:10pt;'>"& prcode &"</td>"
response.write "		</tr>" 
' response.write "		<tr>"
' response.write "			<td class='noborder' style='border-top: 0px solid #111;width:7%; font-size:10pt'>VENDOR</td>"
' response.write "			<td class='noborder' style='border-top: 0px solid #111; font-size:10pt' colspan='6'>: <b>"&spname&"</b></td>"
' response.write "			<td class='sideonly' rowspan='2' style='font-size:10pt;'></td>"
' response.write "			<td rowspan = '2' style='font-size:10pt;'></td>"
' response.write "		</tr>"
' response.write "		<tr>"
' response.write "			<td class='noborder' style='font-size:10pt;'>ADDRESS</td>"
' response.write "			<td class='noborder' colspan='6' style='font-size:10pt;'>: "&spaddress&"</td>"
' response.write "			<td class='sideonly'>&nbsp;</td>"
' response.write "			<td></td>"
' response.write "		</tr>"
' response.write "		<tr>"
' response.write "			<td class='noborder'>&nbsp;</td>"
' response.write "			<td class='noborder' colspan='6'>&nbsp;&nbsp;"&spcity & add_minus & spstate &"</td>"
' response.write "			<td class='sideonly'>&nbsp;</td>"
' response.write "			<td>&nbsp;</td>"
' response.write "		</tr>"
response.write "		<tr>"
response.write "			<td class='noborder' style='font-size:10pt;'>PHONE : </td>"
response.write "			<td class='noborder' colspan='6' style='font-size:10pt;'>: "&phone&" </td>"
response.write "			<td style='font-size:10pt;'></td>"
response.write "			<td style='font-size:10pt;'></td>"
response.write "		</tr>"
response.write "		<tr>"
response.write "			<td class='noborder' style='font-size:10pt;'>FAX</td>"
response.write "			<td class='noborder' colspan='6' style='font-size:10pt;'>: "&fax&"</td>"
response.write "			<td style='font-size:10pt;'></td>"
response.write "			<td style='font-size:10pt;'></td>"
response.write "		</tr>"
' response.write "		<tr>"
' response.write "			<td class='noborder' style='font-size:10pt;'>REKENING PEMBAYARAN</td>"
' response.write "			<td class='noborder' colspan='6' style='font-size:10pt;'>: Bank : "&bankname&", No.Rek. "&bankaccount&" ,A/N "&bankownership&"</td>"
' response.write "			<td style='font-size:10pt;'></td>"
' response.write "			<td style='font-size:10pt;'></td>"
' response.write "		</tr>"
response.write "	</tbody>"
response.write "</table>"
response.write "<div style='padding:7px 0;'></div>"

response.write "<table>"
response.write "	<thead>"
response.write "	</thead>"
response.write "	<tbody>"
' response.write "		<tr>"
' response.write "			<td width='56.6666666667%' align='center' colspan='3'><b>Note</b></td>"
' response.write "			<td width='4.66666666667%' align='center' colspan='3'><b>Transportation</b></td>"
' response.write "			<td width='13.3333333333%' align='center' colspan='3'><b>Delivery Date</b></td>"
' response.write "		</tr>"
' response.write "		<tr>"
' response.write "			<td height='45px' colspan='3'>"&remarks&"</td>"
' response.write "			<td colspan='3'>&nbsp;</td>"
' response.write "			<td colspan='3'>&nbsp;</td>"
' response.write "		</tr>"
response.write "	</tbody>"
response.write "</table>"
response.write "<div style='padding:7px 0;'></div>"
end if
pocode = request("pocode")
q = " SELECT pod.fcba" & _
	"        , pod.prcode" & _
	"        , pod.prlineid" & _
	"        , st.fcname stockname" & _
	"        , pod.quantity" & _
	"        , st.unitofmeasurement uom" & _
	"        , pod.itemcode stockcode" & _
	"        , (pod.quantity * pod.unitprice) - ((pod.amount* nvl(pod.disc,0))/100) amount" & _
	"		 , st.unitofmeasurement" & _
	"		 , po.totalppn" & _
	"		 , po.totalamountppn" & _
	"		 , po.totalcharge" & _
	"		 , po.grandtotal" & _
	"   	 , pod.disc " & _
    "   	 , pod.amountcharge" & _
    "   	 , pod.othercharge" & _
    "   	 , pod.unitprice" & _
	"        , TRIM(REPLACE(REPLACE(REPLACE(pod.otheritemdesc,CHR(10),''),CHR(13),''),CHR(9),'')) otheritemdesc" & _
	" FROM purchaseorderdetails pod" & _
	" INNER JOIN stock st ON (pod.itemcode = st.fccode AND pod.fcba = st.fcba)" & _
	" INNER JOIN purchaseorder po ON (po.pocode = pod.pocode and po.fcba = pod.fcba)" & _
	" WHERE pod.fcba = '" & fcba & "'" & _
	"       AND pod.pocode = '" & pocode & "'"

' response.Write q	
Set oRecordset = Server.CreateObject("ADODB.Recordset")
oRecordset.Open q, connstr

arr_rdata_eof = true
if not oRecordset.eof then
	arr_rdata = oRecordset.getrows()
	arr_rdata_eof = false
End if
oRecordset.close
set oRecordset = Nothing


if not arr_rdata_eof then
	no = 0
	ntotal = 0
	nppn = 0
	gntotal = 0
	response.write "<table border='1'>"
	response.write "	<thead>"
	response.write "		<tr>"
	response.write "			<th style='min-width:2mm;max-width:2mm; font-size:10pt;'>No</th>"
	' response.write "			<th style='min-width:2mm;max-width:2mm;'>PR</th>"
	' response.write "			<th style='min-width:4mm;max-width:4mm;'>Line</th>"
	response.write "			<th style='min-width:100mm;max-widt100mm; font-size:10pt;'>Description</th>"
	response.write "			<th style='min-width:4mm;max-width:4mm; font-size:10pt;'>Qty</th>"
	response.write "			<th style='min-width:10mm;max-width:10mm; font-size:10pt;'>UOM</th>"
	' response.write "			<th style='min-width:5px;max-width:5px;'>Stock Code</th>"
	response.write "			<th style='min-width:10mm;max-width:10mm;font-size:10pt;'>Price</th>"
	' response.write "			<th style='min-width:10mm;max-width:10mm;font-size:10pt;'>Disc (%)</th>"
	response.write "			<th style='min-width:15mm;max-width:15mm;font-size:10pt;'>Total Amount</th>"
	response.write "		</tr>"
	response.write "	</thead>"
	response.write "	<tbody>"
	for yy=0 to ubound(arr_rdata,2)
		no = no +1
		ntotal = ntotal +cdbl(arr_rdata(16,yy))
		response.write "		<tr>"
		response.write "			<td align='center' class='sideonly' style='font-size:10pt;'>"&no&"</td>"
		' response.write "			<td align='center' class='sideonly'>"&arr_rdata(1,yy)&"</td>"
		' response.write "			<td align='center' class='sideonly'>"&arr_rdata(2,yy)&"</td>"
		response.write "			<td class='sideonly' style='font-size:10pt;'>"&arr_rdata(3,yy)&"<br/>"&arr_rdata(17,yy)&"</td>"
		response.write "			<td align='center' class='sideonly' style='font-size:10pt;'>"&formatnumber(cdbl(to_nol(arr_rdata(4,yy))),2)&"</td>"
		response.write "			<td align='center' class='sideonly' style='font-size:10pt;'>"&arr_rdata(5,yy)&"</td>"
		' response.write "			<td align='center' class='sideonly'>"&arr_rdata(16,yy)&"</td>"
		response.write "			<td align='center' class='sideonly' style='font-size:10pt;'>"&FormatNumber(cdbl(to_nol(arr_rdata(16,yy))),0)&"</td>"
		' response.write "			<td align='center' class='sideonly' style='font-size:10pt;'>"&arr_rdata(13,yy)&"</td>"
		' response.write "			<td align='right' class='sideonly'>"&FormatNumber(cdbl(arr_rdata(7,yy))/cdbl(arr_rdata(4,yy)),0)&"</td>"
		response.write "			<td align='right' class='sideonly' style='font-size:10pt;'>"&FormatNumber(cdbl(to_nol(arr_rdata(7,yy))), 0)&"</td>"
		response.write "		</tr>"
	next
	response.write "		<tr>"
	response.write "			<td class='sideonly'>&nbsp;</td>"
	response.write "			<td class='sideonly'>&nbsp;</td>"
	response.write "			<td class='sideonly'>&nbsp;</td>"
	response.write "			<td class='sideonly'>&nbsp;</td>"
	response.write "			<td class='sideonly'>&nbsp;</td>"
	response.write "			<td class='sideonly'>&nbsp;</td>"
	' response.write "			<td class='sideonly'>&nbsp;</td>"
	' response.write "			<td class='sideonly'>&nbsp;</td>"
	' response.write "			<td class='sideonly'>&nbsp;</td>"
	' response.write "			<td class='sideonly'>&nbsp;</td>"
	response.write "		</tr>"
	response.write "		<tr><td colspan='5' align='right' style='border-bottom: none; border-left:none;font-size:10pt;'>Sub Total</td><td align='right' style='font-size:10pt;' >"&FormatNumber(cdbl(to_nol(totalamount)),0)&"</td></tr>"
	' if cstr(""&use_ppn) = "Y" then
	' 	nppn = ntotal * 0.1
	' else
	' 	nppn = 0
	' end if
	' gntotal = ntotal + nppn
	response.write "		<tr><td colspan='5' align='right' class='noborder' colspan='2'>Ppn (10%)</td><td align='right'style='font-size:10pt;' >"&FormatNumber(cdbl(to_nol(totalppn)),0)&"</td></tr>"
	response.write "		<tr><td colspan='5' align='right' class='noborder' colspan='2'>Total Amount + PPn (10%)</td><td align='right' style='font-size:10pt;'>"&FormatNumber(cdbl(to_nol(totalamountppn)),0)&"</td></tr>"
	response.write "		<tr><td colspan='5' align='right' class='noborder' colspan='2'>Total Charge</td><td align='right' style='font-size:10pt;'>"&FormatNumber(cdbl(to_nol(totalcharge)),0)&"</td></tr>"
	response.write "		<tr><td colspan='5' align='right' class='noborder' colspan='2'>Grand Total</td><td align='right' style='font-size:10pt;'>"&FormatNumber(cdbl(to_nol(grandtotal)),0)&"</td></tr>"
	response.write "	</tbody>"
	response.write "</table>"
end if

response.write "<div style='padding:7px 0;'></div>"
response.write "<table>"
response.write "	<thead>"
response.write "	</thead>"
response.write "	<tbody>"
response.write "		<tr>"
response.write "			<td class='noborder' style='font-size:10pt;'>REKENING PEMBAYARAN :</td>"
response.write "		</tr>"
response.write "		<tr>"
response.write "			<td class='noborder' colspan='6' style='font-size:10pt;'>  No.Rek. "&bankaccount&", Bank : "&bankname&",A/N "&bankownership&"</td>"
response.write "			<td style='font-size:10pt;'></td>"
response.write "			<td style='font-size:10pt;'></td>"
response.write "		</tr>"
response.write "		<tr><td width='100%' align='left' style='border:none; white-space: pre;width:100ch;overflow-wrap:normal;font-size:14px;word-wrap: break-word;'>Note :<br/> "&remarks&"</td></tr>"
response.write "	</tbody>"
response.write "</table>"

response.write "<div style='padding:7px 0;'></div>"
response.write "<table>"
response.write "	<thead>"
response.write "	</thead>"
response.write "	<tbody>"
response.write "		<tr>"
response.write "			<td width='24%' align='center' colspan='3' style='font-size:10pt;' ><b><i>Created By</i></b> </td>"
' if cdbl(to_nol(grandtotal)) >= 400000000 then 
response.write "			<td width='24%' align='center' colspan='2' style='font-size:10pt;' ><b><i>Checked By</i></b></td>"
response.write "			<td width='24%' align='center' colspan='2' style='font-size:10pt;' ><b><i>Authorized Signature</i></b></td>"
' else
' response.write "			<td width='24%' align='center' colspan='2' style='font-size:10pt;' ><b><i>Authorized Signature</i></b></td>"
' end if
response.write "			<td width='24%' align='center' colspan='2' style='font-size:10pt;' ><b><i>Vendor Signature</i></b></td>"
response.write "		</tr>"
response.write "		<tr>"
response.write "			<td colspan='3' align='center'><br/><br/><br/><br/><input class='isian' type='text' value=''><p style='font-size:12pt;'><b>PURCHASING</b></p></td>"
' if cdbl(to_nol(grandtotal)) >= 400000000 then 
response.write "			<td colspan='2' align='center'><br/><br/><br/><br/><input class='isian' type='text' value=''><p style='font-size:12pt;'><b>"& divhead_ss &"</b></p></td>"
response.write "			<td colspan='2' align='center'><br/><br/><br/><br/><input class='isian' type='text' value=''><p style='font-size:12pt;'><b>DIRECTOR</b></p></td>"
' else
' response.write "			<td colspan='2' align='center'><br/><br/><br/><br/><input class='isian' type='text' value=''><p style='font-size:12pt;'><b>"& divhead_ss &"</b></p></td>"
' end if
response.write "			<td colspan='2' align='center'><br/><br/><br/><br/><input class='isian' type='text' value=''><p style='font-size:12pt;'><b> "&spname&"</b></p></td>"
response.write "		</tr>"
response.write "	</tbody>"
response.write "</table>"
%>
		</article>
		</section>
	</body>
</html>