<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<link rel="stylesheet" href="/css/fbp-table.css?v=20261010c" type="text/css" />
<script type="text/javascript" src="/js/fbp-table.js?v=20261009c"></script>
<style>
	/* Laporan Contractor Invoice: kop, kartu ringkasan, tabel dengan header lengket */
	body { background: #f3f6f5; }
	#wrapper { max-width: none; }
	.ci-head {
		display: flex; flex-wrap: wrap; align-items: flex-start; justify-content: space-between; gap: 16px;
		padding: 18px 22px; background: #fff; border: 1px solid #e1e7e5; border-radius: 14px;
		border-left: 5px solid #1B7863; box-shadow: 0 1px 3px rgba(23, 38, 42, .06);
	}
	.ci-company { font-size: 12px; font-weight: 700; color: #5d6d69; text-transform: uppercase; letter-spacing: .6px; }
	.ci-ba { margin-top: 2px; font-size: 12px; color: #8a9a96; }
	.ci-title { margin: 6px 0 0; font-size: 24px; font-weight: 700; color: #12584a; letter-spacing: .2px; }
	.ci-chips { display: flex; flex-wrap: wrap; gap: 6px; margin-top: 10px; }
	.ci-chip {
		display: inline-flex; align-items: center; gap: 6px; padding: 3px 10px;
		font-size: 12px; color: #17262a; background: #eef5f3; border: 1px solid #d3e6e0; border-radius: 999px;
	}
	.ci-chip b { font-weight: 600; color: #5d6d69; }
	.ci-actions { display: flex; align-items: center; gap: 8px; }
	.ci-actions input[type="button"] { float: none !important; margin: 0 !important; }
	.ci-print {
		min-height: 36px; padding: 0 14px; font: inherit; font-size: 13px; font-weight: 600;
		color: #1B7863; background: #fff; border: 1px solid #b9d9d0; border-radius: 8px; cursor: pointer;
	}
	.ci-print:hover, .ci-print:focus { background: #e8f3f0; border-color: #1B7863; }
	#ci-summary { grid-template-columns: repeat(auto-fit, minmax(190px, 1fr)); }
	.fbt-card .fbt-card-value.ci-sm { font-size: 19px; }
	.ci-bar { height: 6px; margin-top: 8px; background: #e6ecea; border-radius: 999px; overflow: hidden; }
	.ci-bar span { display: block; height: 100%; background: #1B7863; border-radius: 999px; }

	.ci-panel { background: #fff; border: 1px solid #e1e7e5; border-radius: 14px; padding: 12px 14px 4px; box-shadow: 0 1px 3px rgba(23, 38, 42, .06); }
	.ci-panel .fbt-tools { margin-bottom: 10px; }
	.ci-scroll { overflow: auto; max-height: 70vh; border: 1px solid #e1e7e5; border-radius: 10px; }
	.ci-scroll #product-table { margin: 0; border-collapse: separate; border-spacing: 0; min-width: 100%; }
	/* Baris judul di dalam thead hanya untuk Export Excel */
	#product-table tr.ci-xls, #product-table tr.fbt-toolbar-row { display: none; }
	#product-table thead tr:last-child th {
		position: sticky; top: 0; z-index: 2; white-space: nowrap;
		border-width: 0 1px 0 0; border-color: #2f8a75;
	}
	#product-table td { border-width: 0 0 1px 0; border-color: #e6ecea; white-space: nowrap; }
	#product-table td.ci-wrap { white-space: normal; min-width: 180px; }
	#product-table td.ci-remarks { min-width: 280px; }
	#product-table td.ci-num { text-align: right; }
	#product-table td.ci-code { font-weight: 600; color: #12584a; }
	#product-table td.ci-strong { font-weight: 600; }
	#product-table td.ci-muted { color: #5d6d69; }
	#product-table tfoot td {
		position: sticky; bottom: 0; z-index: 1;
		background: #f1f7f5; font-weight: 700; border-top: 2px solid #1B7863; border-bottom: 0;
	}
	.ci-badge { display: inline-block; padding: 2px 9px; border-radius: 999px; font-size: 11.5px; font-weight: 600; line-height: 1.5; background: #eef1f0; color: #4a5a56; }
	.ci-badge.ok { background: #e2f4ec; color: #136a4b; }
	.ci-badge.warn { background: #fff4d6; color: #8a5a00; }
	.ci-badge.bad { background: #fde8e6; color: #a8322a; }
	.ci-badge.info { background: #e6f0fb; color: #1f5a96; }
	.ci-empty { padding: 28px !important; text-align: center; color: #5d6d69; white-space: normal !important; }
	/* Blok tanda tangan dari Bottom_Report.asp (include bersama; dirapikan hanya di halaman ini) */
	#wrapper ~ table.data {
		width: calc(100% - 48px); margin: 0 24px 24px; padding: 18px 12px 22px; box-sizing: border-box;
		background: #fff; border: 1px solid #e1e7e5; border-radius: 14px; box-shadow: 0 1px 3px rgba(23, 38, 42, .06);
		border-collapse: separate; table-layout: fixed;
	}
	#wrapper ~ table.data td { width: 50%; padding: 0 12px; text-align: center; vertical-align: bottom; background: none; border: 0; }
	#wrapper ~ table.data font { font-family: inherit; font-size: 13px; color: #17262a; background: none; }
	#wrapper ~ table.data tr:nth-child(1) td { height: 0; line-height: 0; font-size: 0; }
	#wrapper ~ table.data tr:nth-child(2) font { font-size: 12px; font-weight: 600; color: #5d6d69; text-transform: uppercase; letter-spacing: .6px; }
	#wrapper ~ table.data tr:nth-child(3) td { height: 64px; }
	#wrapper ~ table.data tr:nth-child(4) font {
		display: inline-block; min-width: 220px; min-height: 18px; padding-top: 6px;
		border-top: 1px solid #8a9a96; font-weight: 600; vertical-align: top;
	}
	@media (max-width: 640px) {
		#wrapper ~ table.data { width: calc(100% - 32px); margin: 0 16px 16px; }
		#wrapper ~ table.data tr:nth-child(4) font { min-width: 120px; }
	}
	@media (max-width: 640px) {
		#wrapper { padding: 12px 16px; }
		.ci-title { font-size: 20px; }
		#ci-summary { grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 8px; }
		#ci-summary .fbt-card { padding: 10px 12px; }
		#ci-summary .fbt-card-value { font-size: 17px; overflow-wrap: anywhere; }
		.ci-panel { padding: 10px 10px 2px; }
	}
	@media print {
		body { background: #fff; }
		.ci-actions { display: none !important; }
		.ci-head, .ci-panel { border: 0; box-shadow: none; padding: 0; }
		.ci-scroll { overflow: visible; max-height: none; border: 0; }
		#product-table thead tr:last-child th, #product-table tfoot td { position: static; }
		.ci-badge { background: none !important; padding: 0; }
		#wrapper ~ table.data { width: 100%; margin: 24px 0 0; border: 0; box-shadow: none; page-break-inside: avoid; }
	}
</style>
<title>ePlaS Reporting - Contractor Invoice </title>
</head>
<body>
<%
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

'---- Escape tanda kutip untuk literal string SQL
Function SqlStr(v)
	SqlStr = Replace(Trim(v & ""), "'", "''")
End Function

'---- Encode output HTML (aman untuk Null)
Function H(v)
	if IsNull(v) then
		H = ""
	else
		H = Server.HTMLEncode(v & "")
	end if
End Function

'---- Parse tanggal format MM/DD/YYYY, kembalikan Empty jika tidak valid
Function ParseMDY(v)
	Dim p
	ParseMDY = Empty
	p = Split(Trim(v & ""), "/")
	if UBound(p) = 2 then
		if IsNumeric(p(0)) and IsNumeric(p(1)) and IsNumeric(p(2)) then
			if CInt(p(0)) >= 1 and CInt(p(0)) <= 12 and CInt(p(1)) >= 1 and CInt(p(1)) <= 31 then
				ParseMDY = DateSerial(CInt(p(2)), CInt(p(0)), CInt(p(1)))
				if Day(ParseMDY) <> CInt(p(1)) then ParseMDY = Empty
			end if
		end if
	end if
End Function

'---- Tanggal ke string MM/DD/YYYY untuk to_date Oracle
Function ToMDY(d)
	ToMDY = Right("0" & Month(d), 2) & "/" & Right("0" & Day(d), 2) & "/" & Year(d)
End Function

'---- Kelas warna badge status dokumen (hanya tampilan)
Function StateClass(v)
	Dim s
	s = LCase(Trim(v & ""))
	StateClass = ""
	if InStr(s, "cancel") > 0 or InStr(s, "reject") > 0 or InStr(s, "void") > 0 or InStr(s, "batal") > 0 then
		StateClass = "bad"
	elseif InStr(s, "post") > 0 or InStr(s, "approv") > 0 or InStr(s, "final") > 0 or InStr(s, "close") > 0 or InStr(s, "release") > 0 then
		StateClass = "ok"
	elseif InStr(s, "progress") > 0 or InStr(s, "draft") > 0 or InStr(s, "new") > 0 or InStr(s, "open") > 0 or InStr(s, "wait") > 0 then
		StateClass = "warn"
	end if
End Function

'---- Chip filter di kop laporan
Function Chip(lbl, v)
	Chip = ""
	if len(trim(v & "")) > 0 then Chip = "<span class=""ci-chip""><b>" & lbl & "</b>" & H(v) & "</span>"
End Function

'---- Parameter filter
invoice_by = Request("invoice_by")
invoicecode = Request("invoicecode")
agreementcode = Request("contractno")
if len(trim(agreementcode)) = 0 then agreementcode = Request("agreementcode")
transdate_from = ParseMDY(Request("CREATIONDATE_From"))
transdate_to = ParseMDY(Request("CREATIONDATE_To"))
businessunit = Request("businessunit")

q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & SqlStr(businessunit) & "' "
set rsBU = connStr.Execute(q)
if not rsBU.eof then
	fcbaname = rsBU("fcName")
	CompanyName = rsBU("FCCOMPANYNAME")
end if
rsBU.close
Set rsBU = Nothing

'---- controljob dihitung sekali per header (inline view), JOBMASTER di-LEFT JOIN
'---- agar invoice tanpa control job tetap tampil
txtSQL = 	"   SELECT cih.invoicecode," & _
			"          cih.invoice_by," & _
			"          cih.transdate," & _
			"          cih.agreementcode," & _
			"          cih.contractorcode," & _
			"          contractor_api.Get_ContractorName(cih.contractorcode,cih.fcba) contractorname," & _
			"          cih.controljob," & _
			"          JB.FCNAME," & _
			"          cih.isretention_invoice," & _
			"          cih.retention," & _
			"          cih.ppn," & _
			"          cih.amountppn," & _
			"          cih.pph," & _
			"          cih.amountpph," & _
			"          cih.rowstate," & _
			"          cih.fcba," & _
			"          cid.jobcode," & _
			"          SUM (cid.totalfinal) totalfinal," & _
			"          MAX (cid.rate) rate," & _
			"          SUM (cid.amount) amount," & _
			"          SUM (cid.amountretention) amountretention," & _
			"          SUM (cid.amountafter_retention) amountafter_retention," & _
			"          cid.remarks," & _
			"          cih.postingflg" & _
			"     FROM (SELECT h.*," & _
			"                  contractor_api.get_controljob(h.contractorcode, h.fcba) controljob" & _
			"             FROM contractorinvoicehdr h" & _
			"            WHERE h.fcba = '" & SqlStr(businessunit) & "') cih" & _
			"          INNER JOIN contractorinvoicedtl cid" & _
			"             ON (    cih.invoicecode = cid.invoicecode" & _
			"                 AND cih.agreementcode = cid.agreementcode" & _
			"                 AND cih.fcba = cid.fcba)" & _
			"          LEFT JOIN JOBMASTER JB" & _
			"             ON (cih.controljob = JB.FCCODE)" & _
			"    WHERE 1 = 1"

if len(trim(invoicecode)) > 0 then
	txtSQL = txtSQL & " and cih.invoicecode = '" & SqlStr(invoicecode) & "'"
end if
if len(trim(agreementcode)) > 0 then
	txtSQL = txtSQL & " and cih.agreementcode = '" & SqlStr(agreementcode) & "'"
end if
if len(trim(invoice_by)) > 0 then
	txtSQL = txtSQL & " and cih.invoice_by = '" & SqlStr(invoice_by) & "'"
end if
if not IsEmpty(transdate_from) then
	txtSQL = txtSQL & " and cih.transdate >= to_date('" & ToMDY(transdate_from) & "','MM/DD/YYYY')"
	txtSelect = " From " & Day(transdate_from) & " " & MonthName(Month(transdate_from)) & " " & Year(transdate_from)
end if
if not IsEmpty(transdate_to) then
	'---- + 1 agar transaksi di hari terakhir (yang punya jam) ikut terambil
	txtSQL = txtSQL & " and cih.transdate < to_date('" & ToMDY(transdate_to) & "','MM/DD/YYYY') + 1"
	txtSelect = txtSelect & "  to " & Day(transdate_to) & " " & MonthName(Month(transdate_to)) & " " & Year(transdate_to)
end if

txtSQL = txtSQL &	"  GROUP BY cih.invoicecode," & _
					"          cih.invoice_by," & _
					"          cih.transdate," & _
					"          cih.agreementcode," & _
					"          cih.contractorcode," & _
					"          cih.controljob," & _
					"          JB.FCNAME," & _
					"          cih.isretention_invoice," & _
					"          cih.retention," & _
					"          cih.ppn," & _
					"          cih.amountppn," & _
					"          cih.pph," & _
					"          cih.amountpph," & _
					"          cih.rowstate," & _
					"          cih.fcba," & _
					"          cid.jobcode," & _
					"          cid.remarks," & _
					"          cih.postingflg" & _
					"  ORDER BY cih.invoicecode, cih.agreementcode, cih.transdate, cid.jobcode"

if len(trim(txtSelect)) = 0 then
	txtSelect = "Selected by All Data"
end if

'response.Write(txtSQL)
Set rsCat = Server.CreateObject("ADODB.Recordset")
	rsCat.CursorType = 3
	rsCat.CursorLocation = 3
	rsCat.Open txtSQL, connstr
%>
		<div id="wrapper">
		<div class="ci-head">
			<div>
				<div class="ci-company">PT. <%=H(CompanyName)%></div>
				<div class="ci-ba"><%=H(fcbaname)%></div>
				<h2 class="ci-title">Contractor Invoice</h2>
				<div class="ci-chips">
					<span class="ci-chip"><b>Periode</b><%if IsEmpty(transdate_from) and IsEmpty(transdate_to) then%>Semua tanggal<%else%><%=H(Trim(txtSelect))%><%end if%></span>
					<%=Chip("Invoice By", invoice_by)%>
					<%=Chip("Agreement", agreementcode)%>
					<%=Chip("Invoice No", invoicecode)%>
				</div>
			</div>
			<div class="ci-actions">
				<button type="button" class="ci-print" onclick="window.print()">Cetak</button>
				<!--#include file="print_excel.asp"-->
			</div>
		</div>
		<div class="fbt-summary" id="ci-summary"></div>
		<div class="ci-panel">
		<div id="ci-tools"></div>
		<div class="ci-scroll">
 <table border="0" width="100%" id="product-table">
         <thead>
         <tr class="ci-xls">
         <th colspan="22" style=" background-color:#FFFFFF; color:#000000; border:none">
			<h4>PT. <%=H(CompanyName)%></h4>
			<h5><%=H(fcbaname)%></h5>
			<h2>CONTRACTOR INVOICE</h2>
		 </th>
			</tr>
				<tr class="ci-xls">
					<td colspan="22" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=H(txtSelect)%></td>
				</tr>
				<tr>
					<th>Invoice No</th>
					<th>Invoice By</th>
					<th>Agreement Code</th>
					<th>Date</th>
					<th>Contractor Code</th>
					<th>Contractor Name</th>
					<th>Contractor Control Job</th>
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
					<th>Posting</th>
				</tr>
                </thead>
               <tbody>
	<%
	v_amountppn = 0
	v_amountpph = 0
	v_totalfinal = 0
	v_amount = 0
	v_amountretention = 0
	v_amountafter_retention = 0
	v_rows = 0
	if Not rsCat.eof then
		dim postingflg, lastInvoice, curInvoice, retFlag, retClass
		lastInvoice = ""
	While Not rsCat.EOF
		if rsCat("postingflg") & "" = "1" then
			postingflg = "YA"
		else
			postingflg = "TIDAK"
		end if

		controljob = rsCat("controljob") & ""
		if len(rsCat("FCNAME") & "") > 0 then controljob = controljob & " : " & rsCat("FCNAME")

		retFlag = Trim(rsCat("isretention_invoice") & "")
		retClass = ""
		if InStr(",y,yes,1,true,ya,", "," & LCase(retFlag) & ",") > 0 then retClass = " info"
		v_rows = v_rows + 1
	%>
		<tr>
			<td class="ci-code"><%=H(rsCat("invoicecode"))%></td>
			<td><%=H(rsCat("invoice_by"))%></td>
			<td><%=H(rsCat("agreementcode"))%></td>
			<td><%=H(rsCat("transdate"))%></td>
			<td class="ci-muted"><%=H(rsCat("contractorcode"))%></td>
			<td class="ci-wrap ci-strong"><%=H(rsCat("contractorname"))%></td>
			<td class="ci-wrap"><%=H(controljob)%></td>
			<td><%if len(retFlag) > 0 then%><span class="ci-badge<%=retClass%>"><%=H(retFlag)%></span><%end if%></td>
			<td class="ci-num"><%=H(rsCat("retention"))%></td>
			<td class="ci-num"><%=H(rsCat("ppn"))%></td>
			<td class="ci-num"><%=formatnumber(cdbl(to_nol(rsCat("amountppn"))),0)%></td>
			<td class="ci-num"><%=H(rsCat("pph"))%></td>
			<td class="ci-num"><%=formatnumber(cdbl(to_nol(rsCat("amountpph"))),0)%></td>
			<td><%if len(rsCat("rowstate") & "") > 0 then%><span class="ci-badge <%=StateClass(rsCat("rowstate"))%>"><%=H(rsCat("rowstate"))%></span><%end if%></td>
			<td><%=H(rsCat("jobcode"))%></td>
			<td class="ci-num"><%=formatnumber(cdbl(to_nol(rsCat("totalfinal"))),0)%></td>
			<td class="ci-num"><%=formatnumber(cdbl(to_nol(rsCat("rate"))),0)%></td>
			<td class="ci-num ci-strong"><%=formatnumber(cdbl(to_nol(rsCat("amount"))),0)%></td>
			<td class="ci-num"><%=formatnumber(cdbl(to_nol(rsCat("amountretention"))),0)%></td>
			<td class="ci-num ci-strong"><%=formatnumber(cdbl(to_nol(rsCat("amountafter_retention"))),0)%></td>
			<td class="ci-wrap ci-remarks"><%=H(rsCat("remarks"))%></td>
			<td><%if postingflg = "YA" then%><span class="ci-badge ok">YA</span><%else%><span class="ci-badge">TIDAK</span><%end if%></td>
		</tr>
	  <%
		'---- PPn & PPh adalah nilai header: dijumlah sekali per invoice,
		'---- bukan per baris jobcode, agar tidak terhitung ganda
		curInvoice = rsCat("invoicecode") & "|" & rsCat("agreementcode")
		if curInvoice <> lastInvoice then
			v_amountppn = v_amountppn + cdbl(to_nol(rsCat("amountppn")))
			v_amountpph = v_amountpph + cdbl(to_nol(rsCat("amountpph")))
			lastInvoice = curInvoice
		end if
		v_totalfinal = v_totalfinal + cdbl(to_nol(rsCat("totalfinal")))
		v_amount = v_amount + cdbl(to_nol(rsCat("amount")))
		v_amountretention = v_amountretention + cdbl(to_nol(rsCat("amountretention")))
		v_amountafter_retention = v_amountafter_retention + cdbl(to_nol(rsCat("amountafter_retention")))
		rsCat.MoveNext
	Wend
	Else
		response.write "<tr><td colspan='22' class='ci-empty'>Tidak ada invoice kontraktor untuk filter ini.</td></tr>"
	End If
	rsCat.close
	Set rsCat = Nothing
	%>
	</tbody>
	<tfoot>
	<tr bgcolor="#F4F4A0">
		<td colspan="10">Total (<span class="fbt-count"><%=v_rows%></span> baris)</td>
		<td class="ci-num" data-sum="10"><%=FormatNumber(v_amountppn,0)%></td>
		<td></td>
		<td class="ci-num" data-sum="12"><%=FormatNumber(v_amountpph,0)%></td>
		<td></td>
		<td></td>
		<td class="ci-num" data-sum="15"><%=FormatNumber(v_totalfinal,0)%></td>
		<td></td>
		<td class="ci-num" data-sum="17"><%=FormatNumber(v_amount,0)%></td>
		<td class="ci-num" data-sum="18"><%=FormatNumber(v_amountretention,0)%></td>
		<td class="ci-num" data-sum="19"><%=FormatNumber(v_amountafter_retention,0)%></td>
		<td></td>
		<td></td>
	</tr>
	</tfoot>
</table>
		</div>
		<div id="ci-foot"></div>
		</div>
		</div>
<script type="text/javascript">
// Tabel interaktif + kartu ringkasan (hanya tampilan; angka dari baris yang dirender server).
// PPn & PPh nilai header invoice: dijumlah sekali per Invoice No + Agreement, sama seperti total server.
window.addEventListener('load', function () {
	var table = document.getElementById('product-table');
	if (!window.FbpTable || !table) return;   // tetap dipasang walau data kosong: toolbar cari & jumlah baris tetap tampil
	var summary = document.getElementById('ci-summary');
	function fmt(n) { return Math.round(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ','); }
	function txt(r, i) { return (r.cells[i].textContent || '').replace(/\s+/g, ' ').trim(); }
	function num(r, i) { return parseFloat(txt(r, i).replace(/,/g, '')) || 0; }
	function card(label, value, note, extra) {
		return '<div class="fbt-card' + (extra || '') + '"><div class="fbt-card-label">' + label + '</div>' +
			'<div class="fbt-card-value' + (String(value).length > 13 ? ' ci-sm' : '') + '">' + value + '</div>' +
			'<div class="fbt-card-note">' + note + '</div></div>';
	}
	FbpTable.patchExport();
	FbpTable.init(table, {
		pageSize: 20,
		onChange: function (rows) {
			var s = { 10: 0, 12: 0, 15: 0, 17: 0, 18: 0, 19: 0 }, inv = {}, nInv = 0, nPosted = 0;
			rows.forEach(function (r) {
				var key = txt(r, 0) + '|' + txt(r, 2);
				if (!inv[key]) {
					inv[key] = 1; nInv++;
					s[10] += num(r, 10); s[12] += num(r, 12);
					if (txt(r, 21) === 'YA') nPosted++;
				}
				s[15] += num(r, 15); s[17] += num(r, 17); s[18] += num(r, 18); s[19] += num(r, 19);
			});
			Array.prototype.forEach.call(table.tFoot.querySelectorAll('[data-sum]'), function (td) {
				td.textContent = fmt(s[td.getAttribute('data-sum')]);
			});
			var pct = nInv ? Math.round(nPosted * 100 / nInv) : 0;
			var retPct = s[17] ? (s[18] * 100 / s[17]).toFixed(1) : '0';
			summary.innerHTML =
				card('Invoice', fmt(nInv), rows.length + ' baris pekerjaan') +
				card('Total amount', fmt(s[17]), 'sebelum retensi &amp; pajak') +
				card('Retensi', fmt(s[18]), retPct + '% dari amount') +
				card('Setelah retensi', fmt(s[19]), 'amount dikurangi retensi') +
				card('PPn / PPh 23', fmt(s[10]) + ' / ' + fmt(s[12]), 'dihitung sekali per invoice') +
				card('Sudah posting', nPosted + ' / ' + nInv, pct + '% invoice terposting<div class="ci-bar"><span style="width:' + pct + '%"></span></div>', ' accent');
		}
	});
	// Toolbar & pager dipindah ke luar area gulir agar tetap terlihat walau tabel lebar
	var tools = table.querySelector('.fbt-toolbar-row .fbt-tools'), foot = table.parentNode.querySelector('.fbt-foot');
	if (tools) { document.getElementById('ci-tools').appendChild(tools); }
	if (foot) { document.getElementById('ci-foot').appendChild(foot); }
});
</script>
<!--#include file="Bottom_Report.asp"-->
