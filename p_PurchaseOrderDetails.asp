<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>ePlaS Reporting - Purchase Order</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<link rel="stylesheet" href="/css/fbp-table.css?v=20261010c" type="text/css" />
<script type="text/javascript" src="/js/fbp-table.js?v=20261009c"></script>
<style>
	/* Laporan Purchase Order: kop, kartu ringkasan, tabel ala DataTables dengan header lengket */
	body { background: #f3f6f5; }
	.po-head {
		display: flex; flex-wrap: wrap; align-items: flex-start; justify-content: space-between; gap: 16px;
		padding: 18px 22px; background: #fff; border: 1px solid #e1e7e5; border-radius: 14px;
		border-left: 5px solid #1B7863; box-shadow: 0 1px 3px rgba(23, 38, 42, .06);
	}
	.po-company { font-size: 12px; font-weight: 700; color: #5d6d69; text-transform: uppercase; letter-spacing: .6px; }
	.po-ba { margin-top: 2px; font-size: 12px; color: #8a9a96; }
	.po-title { margin: 6px 0 0; font-size: 24px; font-weight: 700; color: #12584a; letter-spacing: .2px; }
	.po-chips { display: flex; flex-wrap: wrap; gap: 6px; margin-top: 10px; }
	.po-chip { display: inline-flex; align-items: center; gap: 6px; padding: 3px 10px; font-size: 12px; color: #17262a; background: #eef5f3; border: 1px solid #d3e6e0; border-radius: 999px; }
	.po-chip b { font-weight: 600; color: #5d6d69; }
	.po-actions { display: flex; align-items: center; gap: 8px; }
	.po-actions input[type="button"] { float: none !important; margin: 0 !important; }
	.po-print { min-height: 36px; padding: 0 14px; font: inherit; font-size: 13px; font-weight: 600; color: #1B7863; background: #fff; border: 1px solid #b9d9d0; border-radius: 8px; cursor: pointer; }
	.po-print:hover, .po-print:focus { background: #e8f3f0; border-color: #1B7863; }
	#po-summary { grid-template-columns: repeat(auto-fit, minmax(190px, 1fr)); }
	.fbt-card .fbt-card-value.po-sm { font-size: 19px; }
	.po-mini { display: flex; flex-wrap: wrap; gap: 4px; margin-top: 6px; }

	.po-panel { background: #fff; border: 1px solid #e1e7e5; border-radius: 14px; padding: 12px 14px 4px; box-shadow: 0 1px 3px rgba(23, 38, 42, .06); }
	.po-panel .fbt-tools { margin-bottom: 10px; }
	.po-scroll { overflow: auto; max-height: 70vh; border: 1px solid #e1e7e5; border-radius: 10px; }
	.po-scroll #product-table { margin: 0; border-collapse: separate; border-spacing: 0; min-width: 100%; }
	/* Baris judul di dalam thead hanya untuk Export Excel */
	#product-table tr.po-xls, #product-table tr.fbt-toolbar-row { display: none; }
	#product-table thead tr:last-child th { position: sticky; top: 0; z-index: 2; white-space: nowrap; border-width: 0 1px 0 0; border-color: #2f8a75; }
	#product-table td { border-width: 0 0 1px 0; border-color: #e6ecea; white-space: nowrap; vertical-align: top; }
	#product-table td.po-num { text-align: right; }
	#product-table td.po-code { font-weight: 600; color: #12584a; }
	#product-table td.po-code a { color: #12584a; text-decoration: underline; text-underline-offset: 2px; }
	#product-table td.po-strong { font-weight: 600; }
	#product-table td.po-muted { color: #5d6d69; }
	#product-table td.po-wrap { white-space: normal; min-width: 200px; }
	#product-table td.po-remarks { white-space: normal; min-width: 260px; max-width: 380px; color: #3d4d49; cursor: pointer; }
	/* Catatan PO panjang & berulang per item: tampil 2 baris, klik untuk buka/tutup */
	.po-rm { white-space: pre-line; display: -webkit-box; -webkit-box-orient: vertical; -webkit-line-clamp: 2; line-clamp: 2; overflow: hidden; }
	.po-remarks.open .po-rm { display: block; -webkit-line-clamp: unset; overflow: visible; }
	.po-remarks:not(.open):hover .po-rm { color: #12584a; }
	#product-table tfoot td { position: sticky; bottom: 0; z-index: 1; background: #f1f7f5; font-weight: 700; border-top: 2px solid #1B7863; border-bottom: 0; }
	.po-badge { display: inline-block; padding: 2px 9px; border-radius: 999px; font-size: 11.5px; font-weight: 600; line-height: 1.5; background: #eef1f0; color: #4a5a56; }
	.po-badge.ok { background: #e2f4ec; color: #136a4b; }
	.po-badge.warn { background: #fff4d6; color: #8a5a00; }
	.po-badge.bad { background: #fde8e6; color: #a8322a; }
	.po-badge.info { background: #e6f0fb; color: #1f5a96; }
	.po-empty { padding: 28px !important; text-align: center; color: #5d6d69; white-space: normal !important; }
	@media (max-width: 640px) {
		#wrapper { padding: 12px 16px; }
		.po-title { font-size: 20px; }
		#po-summary { grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 8px; }
		#po-summary .fbt-card { padding: 10px 12px; }
		#po-summary .fbt-card-value { font-size: 17px; overflow-wrap: anywhere; }
		.po-panel { padding: 10px 10px 2px; }
	}
	@media print {
		body { background: #fff; }
		.po-actions { display: none !important; }
		.po-head, .po-panel { border: 0; box-shadow: none; padding: 0; }
		.po-scroll { overflow: visible; max-height: none; border: 0; }
		#product-table thead tr:last-child th, #product-table tfoot td { position: static; }
		.po-badge { background: none !important; padding: 0; }
		.po-rm { display: block; -webkit-line-clamp: unset; overflow: visible; }
	}
</style>
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
<body>
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

cmdremarks = " white-space: pre-line; width: 100ch; overflow-wrap: normal;"

'---- Escape tanda kutip untuk literal string SQL
Function SqlStr(v)
	SqlStr = Replace(Trim(v & ""), "'", "''")
End Function

'---- Encode output HTML (aman untuk Null)
Function H(v)
	If IsNull(v) Then H = "" Else H = Server.HTMLEncode(v & "")
End Function

'---- Tanggal MM/DD/YYYY yang valid, selain itu kosong
Function CleanMDY(v)
	Dim p
	CleanMDY = ""
	p = Split(Trim(v & ""), "/")
	If UBound(p) = 2 Then
		If IsNumeric(p(0)) And IsNumeric(p(1)) And IsNumeric(p(2)) Then
			If IsDate(DateSerial(CInt(p(2)), CInt(p(0)), CInt(p(1)))) Then
				If Day(DateSerial(CInt(p(2)), CInt(p(0)), CInt(p(1)))) = CInt(p(1)) And CInt(p(0)) >= 1 And CInt(p(0)) <= 12 Then CleanMDY = Trim(v & "")
			End If
		End If
	End If
End Function

'---- Kelas warna badge status PO (hanya tampilan)
Function PoStateClass(v)
	Dim s
	s = LCase(Trim(v & ""))
	PoStateClass = ""
	If InStr(s, "cancel") > 0 Or InStr(s, "reject") > 0 Or InStr(s, "batal") > 0 Then
		PoStateClass = "bad"
	ElseIf InStr(s, "order") > 0 Or InStr(s, "approv") > 0 Or InStr(s, "close") > 0 Or InStr(s, "complete") > 0 Then
		PoStateClass = "ok"
	ElseIf InStr(s, "progress") > 0 Or InStr(s, "authoriz") > 0 Or InStr(s, "confirm") > 0 Then
		PoStateClass = "info"
	ElseIf InStr(s, "plan") > 0 Or InStr(s, "draft") > 0 Or InStr(s, "new") > 0 Then
		PoStateClass = "warn"
	End If
End Function

'---- Chip filter di kop laporan
Function Chip(lbl, v)
	Chip = ""
	If Len(Trim(v & "")) > 0 Then Chip = "<span class=""po-chip""><b>" & lbl & "</b>" & H(v) & "</span>"
End Function

'---- Parameter filter: tanda kutip di-escape sebelum disambung ke SQL
v_txtPODATE_From=CleanMDY(request("PODATE_From"))
v_txtPODATE_To=CleanMDY(request("PODATE_To"))
v_txtPOCODE=SqlStr(request("POCODE"))
v_txtITEMCODE=SqlStr(request("ITEMCODE"))
v_txtSUPPLIERCODE=SqlStr(request("SUPPLIERCODE"))
ROWSTATEID	=	SqlStr(request("ROWSTATEID")) 'ditambah ukie tgl 9 okt 2013 untuk filter status
 deliverydestination = SqlStr(request("deliverydestination"))
 purchaseby = SqlStr(request("purchaseby"))

businessunit= SqlStr(request("businessunit"))

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

if request("txtKind") = "POOutstanding" Then
   txtSQL = "  SELECT po.fcba," & _
			"         po.pocode," & _
			"         po.podate," & _
			"         TO_CHAR (po.podate, 'DD') podate_dd," & _
			"         TO_CHAR (po.podate, 'MM') podate_mm," & _
			"         TO_CHAR (po.podate, 'YYYY') podate_YY," & _
			"         po.suppliercode spcode," & _
			"         sp.fcname spname," & _
			"         sp.address spaddress," & _
			"         sp.city spcity," & _
			"         sp.state spstate," & _
			"         po.remarks," & _
			"         po.rowstateid," & _
			"         po.company," & _
			"         po.ppn," & _
			"         po.totalamountppn," & _
			"         po.supplieraddress," & _
			"         po.purchaseby," & _
			"         po.deliverydestination," & _
			"         po.dp," & _
			"         po.totalcharge," & _
			"         po.grandtotal," & _
			"         po.totalamount," & _
			"         po.totalppn," & _
			"         NVL (SUM (NVL (POD.QUANTITY, 0)), 0) AS QUANTITY," & _
			"         NVL (SUM (NVL (POD.RECEIVED, 0)), 0) AS RECEIVED," & _
			"         NVL (SUM (NVL (POD.AMOUNT, 0)), 0) AS AMOUNT," & _
			"         pod.prcode," & _
			"         pod.prlineid," & _
			"         stock_api.get_stockName(pod.itemcode, po.fcba) stockname," & _
			"         st.unitofmeasurement uom," & _
			"         pod.itemcode stockcode," & _
			"         st.unitofmeasurement," & _
			"         po.totalppn," & _
			"         po.totalamountppn," & _
			"         po.totalcharge," & _
			"         po.grandtotal," & _
			"         pod.disc," & _
			"         pod.amountcharge," & _
			"         pod.othercharge," & _
			"         pod.unitprice," & _
			"         TRIM (" & _
			"            REPLACE (" & _
			"               REPLACE (REPLACE (pod.otheritemdesc, CHR (10), '')," & _
			"                        CHR (13)," & _
			"                        '')," & _
			"               CHR (9)," & _
			"               ''))" & _
			"            otheritemdesc," & _
			"         bu2.address addresspurchaseby," & _
			"         bu2.fcname fcnamepurchaseby," & _
			"         bu2.fcfax faxpurchaseby," & _
			"         bu2.fctelephone telephonepurchaseby," & _
			"         bu3.fcfax faxdestination," & _
			"         bu3.fctelephone telephonedestination," & _
			"         bu3.address addressdestination," & _
			"         bu3.fcname fcnamedestination," & _
			"         (SELECT fcname from document_classification where document_type='PO' and fccode=po.document_classification) document_classification" & _
			"    FROM purchaseorder po" & _
			"         INNER JOIN purchaseorderdetails pod" & _
			"            ON (pod.pocode = po.pocode AND pod.fcba = po.fcba)" & _
			"         INNER JOIN stock st" & _
			"            ON (pod.itemcode = st.fccode AND pod.fcba = st.fcba)" & _
			"         LEFT JOIN supplier sp" & _
			"            ON (po.suppliercode = sp.fccode AND po.fcba = sp.fcba)" & _
			"         LEFT JOIN (SELECT *" & _
			"                      FROM businessunit" & _
			"                     WHERE fccode = '"&  purchaseby &"') bu2" & _
			"            ON (bu2.fccode = po.purchaseby)" & _
			"         LEFT JOIN (SELECT *" & _
			"                      FROM businessunit" & _
			"                     WHERE fccode = '"& deliverydestination &"') bu3" & _
			"            ON (bu3.fccode = po.deliverydestination)" & _
			"   WHERE po.purchaseby = '"& SqlStr(Session("fcba")) &"'" & _
		 " and POD.received < POD.QUANTITY"
else
	txtSQL	=	"  SELECT po.fcba," & _
				"         po.pocode," & _
				"         po.podate," & _
				"         TO_CHAR (po.podate, 'DD') podate_dd," & _
				"         TO_CHAR (po.podate, 'MM') podate_mm," & _
				"         TO_CHAR (po.podate, 'YYYY') podate_YY," & _
				"         po.suppliercode spcode," & _
				"         sp.fcname spname," & _
				"         sp.address spaddress," & _
				"         sp.city spcity," & _
				"         sp.state spstate," & _
				"         po.remarks," & _
				"         po.rowstateid," & _
				"         po.company," & _
				"         NVL (SUM (NVL (po.ppn, 0)), 0) AS ppn," & _
				"         po.totalamountppn," & _
				"         po.supplieraddress," & _
				"         po.purchaseby," & _
				"         po.deliverydestination," & _
				"         po.dp," & _
				"         po.totalcharge," & _
				"         po.grandtotal," & _
				"         po.totalamount," & _
				"         po.totalppn," & _
				"         NVL (SUM (NVL (POD.QUANTITY, 0)), 0) AS QUANTITY," & _
				"         NVL (SUM (NVL (POD.UNITPRICE, 0)), 0) AS UNITPRICE," & _
				"         NVL (SUM (NVL (POD.AMOUNT, 0)), 0) AS AMOUNT," & _
				"         pod.prcode," & _
				"         pod.prlineid," & _
				"         stock_api.get_stockName(pod.itemcode, po.fcba) stockname," & _
				"         st.unitofmeasurement uom," & _
				"         pod.itemcode stockcode," & _
				"         st.unitofmeasurement," & _
				"         po.totalppn," & _
				"         po.totalamountppn," & _
				"         po.totalcharge," & _
				"         po.grandtotal," & _
				"         pod.disc," & _
				"         pod.amountcharge," & _
				"         pod.othercharge," & _
				"         pod.UNITPRICE," & _
				"         TRIM (" & _
				"            REPLACE (" & _
				"               REPLACE (REPLACE (pod.otheritemdesc, CHR (10), '')," & _
				"                        CHR (13)," & _
				"                        '')," & _
				"               CHR (9)," & _
				"               ''))" & _
				"            otheritemdesc," & _
				"         bu2.address addresspurchaseby," & _
				"         bu2.fcname fcnamepurchaseby," & _
				"         bu2.fcfax faxpurchaseby," & _
				"         bu2.fctelephone telephonepurchaseby," & _
				"         bu3.fcfax faxdestination," & _
				"         bu3.fctelephone telephonedestination," & _
				"         bu3.address addressdestination," & _
				"         bu3.fcname fcnamedestination ," & _
				"         (SELECT fcname from document_classification where document_type='PO' and fccode=po.document_classification) document_classification" & _
				"    FROM purchaseorder po" & _
				"         INNER JOIN purchaseorderdetails pod" & _
				"            ON (pod.pocode = po.pocode AND pod.fcba = po.fcba)" & _
				"         INNER JOIN stock st" & _
				"            ON (pod.itemcode = st.fccode AND pod.fcba = st.fcba)" & _
				"         LEFT JOIN supplier sp" & _
				"            ON (po.suppliercode = sp.fccode AND po.fcba = sp.fcba)" & _
				"         LEFT JOIN (SELECT *" & _
				"                      FROM businessunit" & _
				"                     WHERE fccode = '"& purchaseby &"') bu2" & _
				"            ON (bu2.fccode = po.purchaseby)" & _
				"         LEFT JOIN (SELECT *" & _
				"                      FROM businessunit" & _
				"                     WHERE fccode = '"& deliverydestination &"') bu3" & _
				"            ON (bu3.fccode = po.deliverydestination)" & _
				"   WHERE PO.podate >=to_Date('" & v_txtPODATE_From & "','MM/DD/YYYY')" & _
				" 	AND PO.podate <=to_Date('" & v_txtPODATE_To & "','MM/DD/YYYY')"

end if		 

if  len(trim(businessunit)) > 0 then
    txtSQL =txtSQL &  " and PO.fcba='" & businessunit & "'"
	txtSelect = txtSelect & " PO No. " & businessunit
end if 
if  len(trim(v_txtPODATE_From)) > 0 then 
	txtSelect = " From " & FormatDateTime(v_txtPODATE_From,1)
end if
if  len(trim(v_txtPODATE_To)) > 0 then 
	txtSelect = txtSelect & "  to " & FormatDateTime(v_txtPODATE_To,1)
end if
if  len(trim(v_txtPOCODE)) > 0 then
    txtSQL = txtSQL & " and PO.pocode='" & v_txtPOCODE & "'"
	txtSelect = txtSelect & " PO Code " & v_txtPOCODE
end if
if  len(trim(v_txtITEMCODE)) > 0 then
    txtSQL =txtSQL &  " and POD.itemcode='" & v_txtITEMCODE & "'"
	txtSelect = txtSelect & " Item Code " & v_txtITEMCODE
end if 
if  len(trim(v_txtSUPPLIERCODE)) > 0 then
    txtSQL =txtSQL &  " and PO.Suppliercode='" & v_txtSUPPLIERCODE & "'"
	txtSelect = txtSelect & " Supplier Code " & v_txtSUPPLIERCODE
end if


'ditambah ukie tgl 9 okt 2013 untuk filter status
If Len(Trim(ROWSTATEID)) > 0 Then
	txtSQL	=	txtSQL &	" and PO.ROWSTATEID = '" & ROWSTATEID & "'"
	txtSelect =	txtSelect &	" Status " & ROWSTATEID
End If 

	txtsql	=	(txtsql)	&	" GROUP BY po.fcba," & _
								"         po.pocode," & _
								"         po.podate," & _
								"         TO_CHAR (po.podate, 'DD')," & _
								"         TO_CHAR (po.podate, 'MM')," & _
								"         TO_CHAR (po.podate, 'YYYY')," & _
								"         po.suppliercode," & _
								"         sp.fcname," & _
								"         sp.address," & _
								"         sp.city," & _
								"         sp.state," & _
								"         po.remarks," & _
								"         po.rowstateid," & _
								"         po.company," & _
								"         po.ppn," & _
								"         po.totalamountppn," & _
								"         po.supplieraddress," & _
								"         po.purchaseby," & _
								"         po.deliverydestination," & _
								"         po.dp," & _
								"         po.totalcharge," & _
								"         po.grandtotal," & _
								"         po.totalamount," & _
								"         po.totalppn," & _
								"         bu2.address," & _
								"         bu2.fcname," & _
								"         bu2.fcfax," & _
								"         bu2.fctelephone," & _
								"         bu3.fcfax," & _
								"         bu3.fctelephone," & _
								"         bu3.address," & _
								"         bu3.fcname," & _
								"         pod.prcode," & _
								"         pod.prlineid," & _
								"        /* st.fcname,*/" & _
								"         st.unitofmeasurement," & _
								"         pod.itemcode," & _
								"         st.unitofmeasurement," & _
								"         po.totalppn," & _
								"         po.totalamountppn," & _
								"         po.totalcharge," & _
								"         po.grandtotal," & _
								"         pod.disc," & _
								"         pod.amountcharge," & _
								"         pod.othercharge," & _
								"         pod.unitprice," & _
								"         pod.otheritemdesc," & _
								" 		  pod.LINENOKEY," & _ 
								"         po.document_classification"

	txtsql	=	(txtsql)	&	" Order by PO.pocode, po.fcba, POD.LINENOKEY"

if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
end if

'response.Write(txtsql)
response.Flush()
%>
 
		<div id="wrapper">
		<%
		Set rsCat = connStr.Execute(txtSQL)
		if Not rsCat.eof then
		   nCountofField=rsCat.fields.count-1
		end if
		NO=1
		if request("txtKind") = "POOutstanding" then
			reportTitle = "Purchase Order Outstanding"
		else
			reportTitle = "Purchase Order"
		end if
		%>
		<div class="po-head">
			<div>
				<%If len(trim(businessunit)) > 0 Then%>
				<div class="po-company">PT. <%=H(CompanyName)%></div>
				<div class="po-ba"><%=H(FcbaName)%></div>
				<%End If%>
				<h2 class="po-title"><%=reportTitle%></h2>
				<div class="po-chips">
					<%if len(v_txtPODATE_From) > 0 or len(v_txtPODATE_To) > 0 then%>
					<span class="po-chip"><b>Periode</b><%=H(v_txtPODATE_From)%> &ndash; <%=H(v_txtPODATE_To)%></span>
					<%elseif request("txtKind") = "POOutstanding" then%>
					<span class="po-chip"><b>Jenis</b>Qty diterima &lt; qty order</span>
					<%else%>
					<span class="po-chip"><b>Periode</b>Semua tanggal</span>
					<%end if%>
					<%=Chip("PO Code", request("POCODE"))%>
					<%=Chip("Item", request("ITEMCODE"))%>
					<%=Chip("Supplier", request("SUPPLIERCODE"))%>
					<%=Chip("Status", request("ROWSTATEID"))%>
				</div>
			</div>
			<div class="po-actions printhide">
				<button type="button" class="po-print" onclick="window.print()">Cetak</button>
				<!--#include file="print_excel.asp"-->
			</div>
		</div>
		<div class="fbt-summary" id="po-summary"></div>
		<div class="po-panel">
		<div id="po-tools"></div>
		<div class="po-scroll">
         <table border="0" width="100%" id="product-table">
         <thead>
         <tr class="po-xls">
         <th colspan="15" style=" background-color:#FFFFFF; color:#000000; border:none">
			<%If len(trim(businessunit)) > 0 Then%>
				<h4>PT. <%=H(CompanyName)%></h4>
				<h5><%=H(FcbaName)%></h5>
			<%End If%>
			<h2><%=UCase(reportTitle)%></h2>
		 </th>
		</tr>
		<tr class="po-xls">
			<th colspan="15" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=H(txtSelect)%></th>
		</tr>
		<tr>
			<th>Date</th>
			<th>Business Unit</th>
			<th>PO.Code</th>
			<th>Classification</th>
			<th>PR.Code</th>
			<th>Supplier</th>
			<th>Status</th><!--'ditambah ukie tgl 9 okt 2013 untuk tampilkan status-->
			<th>Stock Code</th>
			<th>Stock Description</th>
			<th>Remarks</th>
			<th>Order (Qty)</th>
			<th>Price</th>
			<th>Amount</th>
			<th>PPn (%)</th>
			<th>Grand Total</th>
		</tr>
		</thead>
		<tbody>
			<%
			Dim v_Opening,v_Debit,v_Credit,v_Saldo,v_Quantity, v_Received,qtyview, hargaview,ppnhtng, ppnview,amountview,grantotalview
			v_Quantity=0
			v_Received = 0
			v_Amount = 0
			v_grandtotal = 0

			While Not rsCat.EOF

				qtyview 		= cdbl(rscat("Quantity"))
				hargaview 		= cdbl(rscat("UNITPRICE"))
				ppnhtng 		= cdbl(rscat("ppn"))

				'amountview 		= cdbl(qtyview)*cdbl(hargaview)
				amountview 		= cdbl(rscat("Amount"))
				ppnview			= round(cdbl(amountview)*(cdbl(ppnhtng)/100))
				'grantotalview	= cdbl(qtyview)*cdbl(hargaview)+cdbl(ppnview)
				grantotalview	= amountview+ppnview
				'Response.Write ppnhtng

			%>
		<tr data-ppn="<%=ppnview%>">
			<td><%=H(rscat("podate"))%></td>
			<td class="po-muted"><%=H(rscat("fcba"))%></td>
			<%if rsCat("rowstateid") = "Ordered" then %>
			<td class="po-code">
				<a title="Cetak PO" href="javascript:launchwin('printingpo.Asp?txtTable=<%=Server.URLEncode(txtTable & "")%>&txtModule=<%=Server.URLEncode(txtModule & "")%>&pocode=<%=Server.URLEncode(rscat("pocode") & "")%>&podate=<%=Server.URLEncode(rscat("podate") & "")%>&purchaseby=<%=Server.URLEncode(rscat("purchaseby") & "")%>&deliverydestination=<%=Server.URLEncode(rscat("deliverydestination") & "")%>&fcba=<%=Server.URLEncode(rscat("fcba") & "")%>&Mode=Print','newwindow','width='+(980)+',height='+(650)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><%=H(rscat("pocode"))%></a>
			</td>
			<%else%>
			<td class="po-code"><%=H(rscat("pocode"))%></td>
			<%end if%>
			<td><%=H(rscat("document_classification"))%></td>
			<td class="po-muted"><%=H(rscat("prcode"))%></td>
			<td class="po-wrap po-strong"><%=H(rscat("spname"))%></td>
			<!--'ditambah ukie tgl 9 okt 2013 untuk tampilkan status-->
			<td><%if len(rscat("rowstateid") & "") > 0 then%><span class="po-badge <%=PoStateClass(rscat("rowstateid"))%>"><%=H(rscat("rowstateid"))%></span><%end if%></td>
			<td class="po-muted"><%=H(rscat("stockcode"))%></td>
			<td class="po-wrap"><%=H(rscat("stockname"))%></td>
			<td class="po-remarks" title="Klik untuk melihat semua catatan"><div class="po-rm"><%=H(rscat("remarks"))%></div></td>
			<td class="po-num"><%=FormatNumber(cdbl(rscat("Quantity")),2)%></td>
			<td class="po-num"><%=FormatNumber(cdbl(rscat("UNITPRICE")),2)%></td>
			<td class="po-num po-strong"><%=FormatNumber(cdbl(rscat("Amount")),2)%></td>
			<td class="po-num"><%=FormatNumber(cdbl(to_nol(rscat("ppn"))),2)%></td>
			<td class="po-num po-strong"><%=FormatNumber(cdbl(grantotalview),2)%></td>
		</tr>
		  <%
		  v_Quantity=v_Quantity+trim(cdbl(rsCat("Quantity")))
		  'v_Received=v_Received+trim(cdbl(rsCat("Received"))) 'ditambah ukie tgl 9 okt 2013 untuk grandtotal qty terima
		  v_Amount=v_Amount+trim(cdbl(rsCat("Amount")))
		  v_grandtotal = v_grandtotal + trim(cdbl(grantotalview))
		  rsCat.MoveNext
		  no=no+1
		Wend
		rsCat.close()
		Set rsCat = Nothing
		if NO = 1 then%>
		<tr><td colspan="15" class="po-empty">Tidak ada Purchase Order untuk filter ini.</td></tr>
		<%end if%>
		</tbody>
		<tfoot>
		<tr bgcolor="#F4F4A0" style="font-weight: bold;">
			<td colspan="10">Total (<span class="fbt-count"><%=NO-1%></span> baris)</td>
			<td class="po-num" data-sum="10"><%=FormatNumber(v_Quantity,2)%></td>
			<td></td>
			<td class="po-num" data-sum="12"><%=FormatNumber(v_Amount,2)%></td>
			<td></td>
			<td class="po-num" data-sum="14"><%=FormatNumber(v_grandtotal,2)%></td>
		</tr>
		</tfoot>
		</table>
		</div>
		<div id="po-foot"></div>
		</div>
		<input  type="hidden" value="<%=H(txtTable)%>" name="txtTable">
		<input  type="hidden" value="<%=H(txtModule)%>" name="txtModule">
		</div>
<script type="text/javascript">
// Tabel ala DataTables + kartu ringkasan (hanya tampilan; angka dari baris yang dirender server).
window.addEventListener('load', function () {
	var table = document.getElementById('product-table');
	var summary = document.getElementById('po-summary');
	if (!window.FbpTable || !table) return;   // tetap dipasang walau data kosong: toolbar cari & jumlah baris tetap tampil
	function fmt(n, d) { return n.toFixed(d || 0).replace(/\B(?=(\d{3})+(?!\d))/g, ','); }
	function txt(r, i) { return (r.cells[i].textContent || '').replace(/\s+/g, ' ').trim(); }
	function num(r, i) { return parseFloat(txt(r, i).replace(/,/g, '')) || 0; }
	function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
	function card(label, value, note, extra) {
		return '<div class="fbt-card' + (extra || '') + '"><div class="fbt-card-label">' + label + '</div>' +
			'<div class="fbt-card-value' + (String(value).length > 13 ? ' po-sm' : '') + '">' + value + '</div>' +
			'<div class="fbt-card-note">' + note + '</div></div>';
	}
	FbpTable.patchExport();
	FbpTable.init(table, {
		pageSize: 20,
		onChange: function (rows) {
			var s = { 10: 0, 12: 0, 14: 0 }, ppn = 0, po = {}, nPo = 0, sup = {}, nSup = 0, st = {};
			rows.forEach(function (r) {
				s[10] += num(r, 10); s[12] += num(r, 12); s[14] += num(r, 14);
				ppn += parseFloat(r.getAttribute('data-ppn')) || 0;
				var k = txt(r, 1) + '|' + txt(r, 2);
				if (!po[k]) { po[k] = txt(r, 6) || '-'; nPo++; st[po[k]] = (st[po[k]] || 0) + 1; }
				var sp = txt(r, 5);
				if (sp && !sup[sp]) { sup[sp] = 1; nSup++; }
			});
			Array.prototype.forEach.call(table.tFoot.querySelectorAll('[data-sum]'), function (td) {
				td.textContent = fmt(s[td.getAttribute('data-sum')], 2);
			});
			var badges = Object.keys(st).sort(function (a, b) { return st[b] - st[a]; }).slice(0, 4).map(function (k) {
				var c = /cancel|reject/i.test(k) ? 'bad' : /order|approv|close|complete/i.test(k) ? 'ok' : /progress|authoriz|confirm/i.test(k) ? 'info' : /plan|draft|new/i.test(k) ? 'warn' : '';
				return '<span class="po-badge ' + c + '">' + esc(k) + ' ' + st[k] + '</span>';
			}).join('');
			summary.innerHTML =
				card('Purchase Order', fmt(nPo), rows.length + ' baris item &middot; ' + nSup + ' supplier') +
				card('Total amount', fmt(s[12]), 'sebelum PPn') +
				card('PPn', fmt(ppn), 'amount &times; tarif PPn per baris') +
				card('Grand total', fmt(s[14]), 'amount + PPn', ' accent') +
				card('Total qty order', fmt(s[10], 2), 'semua satuan digabung') +
				card('Status PO', nPo ? fmt(nPo) + ' PO' : '-', '<div class="po-mini">' + badges + '</div>');
		}
	});
	table.tBodies[0].addEventListener('click', function (e) {
		var td = e.target.closest ? e.target.closest('td.po-remarks') : null;
		if (td && !e.target.closest('a')) td.classList.toggle('open');
	});
	// Toolbar & pager di luar area gulir agar tetap terlihat walau tabel lebar
	var tools = table.querySelector('.fbt-toolbar-row .fbt-tools'), foot = table.parentNode.querySelector('.fbt-foot');
	if (tools) document.getElementById('po-tools').appendChild(tools);
	if (foot) document.getElementById('po-foot').appendChild(foot);
});
</script>
			<!--#include file="Bottom_Report.asp"-->
