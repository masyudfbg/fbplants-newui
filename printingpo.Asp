<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<!--#include virtual="/email/qrcode.inc"-->
<!--#include virtual="/email/po_link_token.inc"-->
<%
' Link PO dari email supplier: boleh dibuka tanpa login hanya bila token cocok dengan pocode & fcba
If Session("USERNAME") & "" = "" And Request("t") & "" <> "" Then
	If PoLink_Valid(Request("pocode"), Request("fcba"), Request("t")) Then FbpPublicAllowed = True
End If
%>
<%validatesession%>
<%
'---- Escape tanda kutip untuk literal string SQL
Function PpSql(v)
	PpSql = Replace(Trim(v & ""), "'", "''")
End Function

'---- Encode output HTML (aman untuk Null)
Function PpH(v)
	If IsNull(v) Then PpH = "" Else PpH = Server.HTMLEncode(v & "")
End Function

'---- Logo perusahaan per company code (sama dengan cetakan lama & PDF email)
Function PpLogo(code)
	Select Case UCase(Trim(code & ""))
		Case "PT.ARJ" : PpLogo = "logoarj.jpeg"
		Case "PT.BRS" : PpLogo = "logobrs.jpg"
		Case "PT.BIA" : PpLogo = "logobia.jpeg"
		Case "PT.KAA", "POM" : PpLogo = "logokaa.jpg"
		Case "PT.CPP" : PpLogo = "logocpp.png"
		Case "PT.GAS" : PpLogo = "logogas.jpeg"
		Case "PT.KSR" : PpLogo = "logoksr.jpg"
		Case "PT.BRP" : PpLogo = "logobrp.jpg"
		Case "PT.CAE" : PpLogo = "logocae.jpg"
		Case Else : PpLogo = ""
	End Select
End Function

'---- QR code nomor PO sebagai SVG (encoder yang sama dengan PDF lampiran email)
Function PpQrSvg(text)
	Dim mods, size, x, y, d
	PpQrSvg = ""
	mods = PoQr_Encode(text & "", size)
	If IsEmpty(mods) Then Exit Function
	d = ""
	For y = 0 To size - 1
		For x = 0 To size - 1
			If mods(y * size + x) Then d = d & "M" & (x + 2) & " " & (y + 2) & "h1v1h-1z"
		Next
	Next
	PpQrSvg = "<svg class=""po-qr-svg"" xmlns=""http://www.w3.org/2000/svg"" viewBox=""0 0 " & (size + 4) & " " & (size + 4) & """ shape-rendering=""crispEdges"" role=""img"" aria-label=""QR code nomor PO " & PpH(text) & """>" & _
	          "<rect width=""100%"" height=""100%"" fill=""#fff""/><path d=""" & d & """ fill=""#000""/></svg>"
End Function

'---- Satu baris ringkasan biaya (label rata kanan + nilai berbingkai)
Sub PpSumRow(label, value, cls)
	Response.Write "<tr class=""po-sum " & cls & """><td colspan=""5"" class=""po-sum-label"">" & label & "</td><td class=""po-sum-val"">" & FormatNumber(cdbl(to_nol(value)), 0) & "</td></tr>" & vbCrLf
End Sub

 fcba = Request("fcba")
 pocode = request("pocode")
 deliverydestination = request("deliverydestination")
 purchaseby = request("purchaseby")
 		q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & PpSql(fcba) & "' "
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
	"       bu3.fcname fcnamedestination," & _
	"       (SELECT fcname from document_classification where document_type='PO' and fccode=po.document_classification) document_classification" & _
	"  FROM purchaseorderdetails pod" & _
	"		INNER JOIN purchaseorder po ON (pod.pocode = po.pocode AND pod.fcba = po.fcba)" & _
	"       LEFT JOIN supplier sp" & _
	"          ON (po.suppliercode = sp.fccode AND po.fcba = sp.fcba)" & _
	"       LEFT JOIN (SELECT *" & _
	"                    FROM businessunit" & _
	"                   WHERE fccode = '"& PpSql(purchaseby) &"') bu2" & _
	"          ON (bu2.fccode = po.purchaseby)" & _
	"       LEFT JOIN (SELECT *" & _
	"                    FROM businessunit" & _
	"                   WHERE fccode = '"& PpSql(deliverydestination) &"') bu3" & _
	"          ON (bu3.fccode = po.deliverydestination)" & _
	" WHERE po.fcba = '"& PpSql(fcba) &"' AND po.pocode = '"& PpSql(pocode) &"'"

' response.Write q
Set oRecordset = Server.CreateObject("ADODB.Recordset")
oRecordset.Open q, connstr
hdr_found = false
if not oRecordset.eof then
	hdr_found = true
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
	document_classification = oRecordset("document_classification")
End if
oRecordset.close
set oRecordset = Nothing

pocode = request("pocode")
q = " SELECT pod.fcba" & _
	"        , pod.prcode" & _
	"        , pod.prlineid" & _
	"        , stock_api.get_stockName(pod.itemcode, pod.fcba) stockname" & _
	"        , pod.quantity" & _
	"        , st.unitofmeasurement uom" & _
	"        , pod.itemcode stockcode" & _
	"        , (pod.quantity * pod.unitprice) - ((pod.quantity * pod.unitprice) * (nvl(pod.disc,0))/100)  amount" & _
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
	"   	 , pod.otcstock1 oc_pbbkb75" & _
	"   	 , pod.otcstock3 oc_transport" & _
	"   	 , pod.otherchargestock " & _
	" FROM purchaseorderdetails pod" & _
	" INNER JOIN bankstock st ON (pod.itemcode = st.fccode)" & _
	" INNER JOIN purchaseorder po ON (po.pocode = pod.pocode and po.fcba = pod.fcba)" & _
	" WHERE pod.fcba = '" & PpSql(fcba) & "'" & _
	"       AND pod.pocode = '" & PpSql(pocode) & "'"

 'response.Write q
Set oRecordset = Server.CreateObject("ADODB.Recordset")
oRecordset.Open q, connstr

arr_rdata_eof = true
if not oRecordset.eof then
	arr_rdata = oRecordset.getrows()
	arr_rdata_eof = false
End if
oRecordset.close
set oRecordset = Nothing

'---- Nilai tampilan
companyTitle = Trim(CompanyName & "")
if UCase(Left(companyTitle, 3)) <> "PT." and UCase(Left(companyTitle, 3)) <> "PT " then companyTitle = "PT. " & companyTitle
logoFile = PpLogo(CompanyCode)
gudang = Trim(fcnamedestination & "")
if gudang = "" then gudang = "Kebun"
tglPO = ""
if podate_dd & "" <> "" then tglPO = podate_dd & " " & bulan(podate_mm) & " " & podate_yy
%>
<!DOCTYPE html>
<html lang="id">
	<head>
		<meta charset="utf-8">
		<meta name="viewport" content="width=device-width, initial-scale=1">
		<title>Purchase Order <%=PpH(pocode)%></title>
		<link rel="stylesheet" href="/plugins/paper-css/paper.min.css">
		<style>
			/* Cetakan PO: tata letak & font mengikuti PDF lampiran email (Helvetica/Arial) */
			@page { size: A4; margin: 0; }
			html, body { margin: 0; }
			body.A4 { background: #e6eaea; }
			.po, .po * { box-sizing: border-box; }
			.po {
				font-family: Helvetica, Arial, "Liberation Sans", sans-serif;
				font-size: 8pt; line-height: 1.3; color: #000;
			}
			/* Lembar A4 didefinisikan sendiri (paper.css tidak selalu tersedia di server) */
			body.A4 .sheet.po { width: 210mm; height: auto; min-height: 296mm; overflow: visible; padding: 9mm 10mm 10mm; background: #fff; position: relative; }
			.po p { margin: 0; }
			.po b, .po strong { font-weight: 700; }

			/* Tombol (layar saja) */
			.po-toolbar { display: flex; justify-content: flex-end; gap: 8px; margin-bottom: 8px; }
			.po-toolbar input[type="button"], .po-toolbar button {
				float: none !important; margin: 0 !important; display: inline-flex; align-items: center;
				min-height: 34px; padding: 0 14px !important; border-radius: 6px !important;
				font: 600 12px Helvetica, Arial, sans-serif !important; cursor: pointer;
			}
			.po-btn-print { color: #1B7863; background: #fff; border: 1px solid #1B7863; }
			.po-toolbar input[onclick*="tableToExcel"] { color: #fff !important; background: #1B7863 !important; border: 1px solid #1B7863 !important; }

			/* Kop: logo kiri, QR kanan */
			.po-top { display: flex; justify-content: space-between; align-items: flex-start; gap: 12px; min-height: 26mm; }
			.po-logo img { display: block; height: 20mm; max-width: 82mm; object-fit: contain; }
			.po-qr { width: 21mm; text-align: center; flex: 0 0 auto; }
			.po-qr-svg { display: block; width: 21mm; height: 21mm; }
			.po-qr small { display: block; font-size: 5.5pt; margin-top: 1px; }

			/* Dua kolom: kiri 62%, kanan 38% (sama dengan PDF) */
			.po-row { display: grid; grid-template-columns: 62% 38%; }
			.po-company { font-size: 9pt; font-weight: 700; }
			.po-title { font-size: 9pt; font-weight: 700; margin-bottom: 3px; }
			.po-kv { display: grid; grid-template-columns: 26mm 1fr; column-gap: 0; row-gap: 1px; }
			.po-kv.left { grid-template-columns: 17mm 1fr; padding-right: 8mm; }
			.po-kv > span:nth-child(odd) { white-space: nowrap; }
			.po-addr { margin-top: 8px; padding-right: 8mm; }
			.po-addr p + p { margin-top: 6px; }
			.po-info { margin-top: 12px; }

			/* Tabel item */
			.po-scroll { margin-top: 8px; }
			.po-items { width: 100%; border-collapse: collapse; font-size: 8pt; }
			.po-items th { padding: 3px 4px; font-weight: 700; text-align: center; border: 1px solid #000; }
			.po-items td { padding: 3px 4px; vertical-align: top; }
			.po-items tbody td { border-left: 1px solid #000; border-right: 1px solid #000; }
			.po-items tbody tr:last-child td { padding-bottom: 8px; border-bottom: 1px solid #000; }
			.po-items .c-no { width: 5%; text-align: center; }
			.po-items .c-qty { width: 10%; text-align: center; }
			.po-items .c-uom { width: 9%; text-align: center; }
			.po-items .c-price { width: 13%; text-align: center; }
			.po-items .c-amt { width: 15%; text-align: right; }
			.po-items .c-desc span { display: block; }
			.po-sum td { padding: 2px 4px; }
			.po-sum-label { text-align: right; border: 0; }
			.po-sum-val { text-align: right; border: 1px solid #000; }
			.po-sum.first td { border-top: 1px solid #000; border-bottom: 1px solid #000; }
			.po-sum.first .po-sum-label { border-left: 1px solid #000; }

			/* Rekening, catatan, tanda tangan */
			.po-block { margin-top: 16px; }
			.po-block h4 { margin: 0 0 2px; font-size: 8pt; font-weight: 700; }
			.po-block .indent { padding-left: 4px; }
			.po-remarks { white-space: pre-line; overflow-wrap: anywhere; }
			.po-sign { display: grid; grid-template-columns: repeat(4, 1fr); margin-top: 18px; text-align: center; page-break-inside: avoid; }
			.po-sign .lbl { font-weight: 700; font-style: italic; }
			.po-sign .name { margin-top: 62px; font-weight: 700; overflow-wrap: anywhere; padding: 0 4px; }
			.po-empty { padding: 30px 0; text-align: center; color: #555; }

			@media screen {
				body.A4 .sheet.po { margin: 8mm auto; box-shadow: 0 2px 10px rgba(0, 0, 0, .18); }
			}
			@media print {
				body.A4 { background: #fff; }
				.po-toolbar { display: none !important; }
				body.A4 .sheet.po { margin: 0; box-shadow: none; }
				.po-items thead { display: table-header-group; }
				.po-items tr { page-break-inside: avoid; }
			}

			/* Tampilan HP */
			@media screen and (max-width: 820px) {
				body.A4:not(.po-fit) { background: #fff; }
				body.A4:not(.po-fit) .sheet.po { width: auto; min-height: 0; margin: 0; padding: 14px 16px 24px; box-shadow: none; }
				body:not(.po-fit) .po { font-size: 12.5px; line-height: 1.45; }
				body:not(.po-fit) .po-company, body:not(.po-fit) .po-title { font-size: 14px; }
				body:not(.po-fit) .po-toolbar { position: sticky; top: 0; z-index: 5; margin: -14px -16px 12px; padding: 10px 16px; background: #fff; border-bottom: 1px solid #e1e7e5; }
				body:not(.po-fit) .po-top { min-height: 0; margin-bottom: 12px; }
				body:not(.po-fit) .po-logo img { height: 46px; max-width: 60vw; }
				body:not(.po-fit) .po-qr, body:not(.po-fit) .po-qr-svg { width: 72px; }
				body:not(.po-fit) .po-qr-svg { height: 72px; }
				body:not(.po-fit) .po-qr small { font-size: 10px; }
				body:not(.po-fit) .po-row { grid-template-columns: 1fr; }
				body:not(.po-fit) .po-row > .po-right { order: -1; margin-bottom: 10px; padding: 10px 12px; background: #f3f6f5; border-radius: 8px; }
				body:not(.po-fit) .po-kv, body:not(.po-fit) .po-kv.left { grid-template-columns: 96px 1fr; padding-right: 0; row-gap: 3px; }
				body:not(.po-fit) .po-info .po-left { margin-top: 0; }
				body:not(.po-fit) .po-addr { padding-right: 0; }
				body:not(.po-fit) .po-scroll { margin-top: 14px; }
				/* Item jadi kartu per baris: deskripsi di atas, qty/uom/harga & total di bawah */
				body:not(.po-fit) .po-items, body:not(.po-fit) .po-items tbody, body:not(.po-fit) .po-items tfoot { display: block; width: 100%; font-size: 12.5px; }
				body:not(.po-fit) .po-items thead { display: none; }
				body:not(.po-fit) .po-items tbody { border-top: 2px solid #17262a; }
				body:not(.po-fit) .po-items tbody tr { display: grid; grid-template-columns: 24px 60px 40px 1fr auto; column-gap: 6px; row-gap: 2px; padding: 9px 0; border-bottom: 1px solid #dfe6e4; }
				body:not(.po-fit) .po-items tbody td, body:not(.po-fit) .po-items tbody tr:last-child td { display: block; width: auto !important; padding: 0; border: 0; text-align: left; white-space: nowrap; }
				body:not(.po-fit) .po-items tbody td.c-desc { white-space: normal; }
				body:not(.po-fit) .po-items tbody tr:last-child { border-bottom: 2px solid #17262a; }
				body:not(.po-fit) .po-items td.c-no, body:not(.po-fit) .po-items tbody tr:last-child td.c-no { grid-row: 1 / span 2; color: #5d6d69; text-align: left; }
				body:not(.po-fit) .po-items td.c-desc { grid-column: 2 / -1; font-weight: 600; }
				body:not(.po-fit) .po-items td.c-desc span + span { font-weight: 400; color: #5d6d69; }
				body:not(.po-fit) .po .po-items tbody td.c-qty, body:not(.po-fit) .po .po-items tbody td.c-uom, body:not(.po-fit) .po .po-items tbody td.c-price { grid-row: 2; color: #5d6d69; text-align: left; }
				body:not(.po-fit) .po-items td.c-price::before { content: "\00d7  "; }
				body:not(.po-fit) .po-items td.c-amt { grid-row: 2; grid-column: 5; font-weight: 700; text-align: right; }
				body:not(.po-fit) .po-items tfoot { margin-top: 6px; }
				body:not(.po-fit) .po-items tfoot tr { display: flex; justify-content: space-between; gap: 12px; padding: 3px 0; }
				body:not(.po-fit) .po-items tfoot td, body:not(.po-fit) .po-sum.first td { display: block; padding: 0; border: 0 !important; }
				body:not(.po-fit) .po-items tfoot .po-sum-label { text-align: left; color: #5d6d69; }
				body:not(.po-fit) .po-items tfoot tr:last-child { margin-top: 4px; padding-top: 7px; border-top: 1px solid #17262a; font-size: 14px; font-weight: 700; }
				body:not(.po-fit) .po-items tfoot tr:last-child .po-sum-label { color: #17262a; }
				body:not(.po-fit) .po-block h4 { font-size: 13px; }
				body:not(.po-fit) .po-sign { grid-template-columns: 1fr 1fr; row-gap: 22px; }
				body:not(.po-fit) .po-sign .name { margin-top: 44px; }
			}
		</style>
	</head>
	<body class="A4">
		<section class="sheet po">
		<div class="po-toolbar">
			<button type="button" class="po-btn-print" onclick="window.print()">Cetak</button>
			<!--#include file="print_excel.asp"-->
		</div>
		<article id="product-table">
<%if not hdr_found then%>
			<div class="po-empty">Purchase Order <b><%=PpH(pocode)%></b> tidak ditemukan.</div>
<%else%>
			<div class="po-top">
				<div class="po-logo"><%if logoFile <> "" then%><img src="css/login/images/<%=logoFile%>" alt="Logo <%=PpH(companyTitle)%>"><%end if%></div>
				<div class="po-qr"><%=PpQrSvg(pocode)%><small>Scan: No. PO</small></div>
			</div>

			<div class="po-row po-head">
				<div class="po-left">
					<div class="po-company"><%=PpH(companyTitle)%></div>
					<div class="po-addr">
						<p><%=PpH(addresspurchaseby)%></p>
						<p>Phone. <%=PpH(telephonepurchaseby)%><br>Fax. <%=PpH(faxpurchaseby)%></p>
					</div>
				</div>
				<div class="po-right">
					<div class="po-title">PURCHASE ORDER</div>
					<div class="po-kv">
						<span>No :</span><span><%=PpH(pocode)%></span>
						<span>Classification :</span><span><%=PpH(document_classification)%></span>
					</div>
				</div>
			</div>

			<div class="po-row po-info">
				<div class="po-left">
					<div class="po-kv left">
						<span>Kepada</span><span>: <b><%=PpH(spname)%></b></span>
						<span>Alamat</span><span>: <%=PpH(spaddress)%></span>
						<span>Phone</span><span>: <%=PpH(phone)%></span>
						<span>Fax</span><span>: <%=PpH(fax)%></span>
						<span>U/P</span><span>: </span>
					</div>
				</div>
				<div class="po-right">
					<div class="po-kv">
						<span>Tanggal</span><span><%=PpH(tglPO)%></span>
						<span>Berlaku S/D</span><span></span>
						<span>Mata Uang :</span><span>Rupiah</span>
						<span>Gudang :</span><span><%=PpH(gudang)%></span>
						<span>No. PR :</span><span><%=PpH(prcode)%></span>
					</div>
				</div>
			</div>

<%if not arr_rdata_eof then%>
			<div class="po-scroll">
			<table class="po-items">
				<thead>
					<tr>
						<th class="c-no">No</th>
						<th class="c-desc">Description</th>
						<th class="c-qty">Qty</th>
						<th class="c-uom">UOM</th>
						<th class="c-price">Price</th>
						<th class="c-amt">Total Amount</th>
					</tr>
				</thead>
				<tbody>
<%
	no = 0
	for yy = 0 to ubound(arr_rdata, 2)
		no = no + 1
%>
					<tr>
						<td class="c-no"><%=no%></td>
						<td class="c-desc"><span><%=PpH(arr_rdata(3, yy))%></span><%if Trim(arr_rdata(17, yy) & "") <> "" then%><span><%=PpH(arr_rdata(17, yy))%></span><%end if%></td>
						<td class="c-qty"><%=formatnumber(cdbl(to_nol(arr_rdata(4, yy))), 2)%></td>
						<td class="c-uom"><%=PpH(arr_rdata(5, yy))%></td>
						<td class="c-price"><%=FormatNumber(cdbl(to_nol(arr_rdata(16, yy))), 0)%></td>
						<td class="c-amt"><%=FormatNumber(cdbl(to_nol(arr_rdata(7, yy))), 0)%></td>
					</tr>
<%
	next
%>
				</tbody>
				<tfoot>
<%
	Call PpSumRow("Sub Total", totalamount, "first")
	Call PpSumRow("Ppn " & PpH(ppn) & " %", totalppn, "")
	Call PpSumRow("Total Amount + PPn " & PpH(ppn) & " %", totalamountppn, "")
	if cdbl(to_nol(arr_rdata(18, 0))) > 0 then Call PpSumRow("PBBKP (7.5%)", arr_rdata(18, 0), "")
	if cdbl(to_nol(arr_rdata(19, 0))) > 0 then Call PpSumRow("OAT", arr_rdata(19, 0), "")
	if cdbl(to_nol(arr_rdata(15, 0))) > 0 then Call PpSumRow("Other Charge", arr_rdata(15, 0), "")
	if cdbl(to_nol(arr_rdata(20, 0))) > 0 then Call PpSumRow("Other Charge", arr_rdata(20, 0), "")
	Call PpSumRow("Grand Total", grandtotal, "")
%>
				</tfoot>
			</table>
			</div>
<%end if%>

			<div class="po-block">
				<h4>Rekening Pembayaran :</h4>
				<p class="indent">No.Rek. <%=PpH(bankaccount)%>, Bank : <%=PpH(bankname)%>,A/N <%=PpH(bankownership)%></p>
			</div>
			<div class="po-block" style="margin-top:8px">
				<h4>Catatan :</h4>
				<p class="indent po-remarks"><%=PpH(remarks)%></p>
			</div>

			<div class="po-sign">
				<div><div class="lbl">Dibuat Oleh,</div><div class="name">PURCHASING</div></div>
				<div><div class="lbl">Disetujui Oleh,</div><div class="name">PURCHASING DEPT</div></div>
				<div><div class="lbl">Disetujui Oleh,</div><div class="name">HEAD COMM &amp; PURCH DEPT</div></div>
				<div><div class="lbl">Disetujui Supplier,</div><div class="name"><%=PpH(spname)%></div></div>
			</div>
<%end if%>
		</article>
		</section>
		<script>
		// Cetak satu halaman A4: bila isi lebih tinggi dari lembar, isi diperkecil secukupnya
		// (lebar tetap penuh). Tampilan layar tidak berubah. Batas perkecil 55% agar tetap terbaca;
		// PO yang sangat panjang tetap lanjut ke halaman berikutnya.
		(function () {
			var art = document.getElementById('product-table');
			if (!art) return;
			var MM = 96 / 25.4, PAGE_H = 297, PAD_V = 19, BODY_W = 190, MIN_Z = 0.55;
			function reset() { art.style.zoom = ''; art.style.width = ''; document.body.classList.remove('po-fit'); }
			function fit() {
				reset();
				document.body.classList.add('po-fit'); // ukur dengan tata letak A4 walau jendela sempit (tata letak HP dimatikan)
				var need = art.getBoundingClientRect().height / MM + PAD_V, avail = PAGE_H - PAD_V - 2;
				if (need - PAD_V <= avail) return;
				var z = Math.max(MIN_Z, avail / (need - PAD_V));
				art.style.width = (BODY_W / z) + 'mm';
				art.style.zoom = z;
			}
			window.addEventListener('beforeprint', fit);
			window.addEventListener('afterprint', reset);
			if (window.matchMedia) {
				var mq = window.matchMedia('print');
				var h = function (e) { if (e.matches) fit(); else reset(); };
				if (mq.addEventListener) mq.addEventListener('change', h); else if (mq.addListener) mq.addListener(h);
			}
		})();
		</script>
	</body>
</html>
