<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>ePlaS Reporting - Purchase Request</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<link rel="stylesheet" href="/css/fbp-table.css?v=20261010c" type="text/css" />
<script type="text/javascript" src="/js/fbp-table.js?v=20261009c"></script>
<style>
	/* Laporan Purchase Request: kop, kartu ringkasan, tabel ala DataTables dengan header lengket */
	body { background: #f3f6f5; }
	.pr-head {
		display: flex; flex-wrap: wrap; align-items: flex-start; justify-content: space-between; gap: 16px;
		padding: 18px 22px; background: #fff; border: 1px solid #e1e7e5; border-radius: 14px;
		border-left: 5px solid #1B7863; box-shadow: 0 1px 3px rgba(23, 38, 42, .06);
	}
	.pr-company { font-size: 12px; font-weight: 700; color: #5d6d69; text-transform: uppercase; letter-spacing: .6px; }
	.pr-ba { margin-top: 2px; font-size: 12px; color: #8a9a96; }
	.pr-title { margin: 6px 0 0; font-size: 24px; font-weight: 700; color: #12584a; letter-spacing: .2px; }
	.pr-chips { display: flex; flex-wrap: wrap; gap: 6px; margin-top: 10px; }
	.pr-chip { display: inline-flex; align-items: center; gap: 6px; padding: 3px 10px; font-size: 12px; color: #17262a; background: #eef5f3; border: 1px solid #d3e6e0; border-radius: 999px; }
	.pr-chip b { font-weight: 600; color: #5d6d69; }
	.pr-actions { display: flex; align-items: center; gap: 8px; }
	.pr-actions input[type="button"] { float: none !important; margin: 0 !important; }
	.pr-print { min-height: 36px; padding: 0 14px; font: inherit; font-size: 13px; font-weight: 600; color: #1B7863; background: #fff; border: 1px solid #b9d9d0; border-radius: 8px; cursor: pointer; }
	.pr-print:hover, .pr-print:focus { background: #e8f3f0; border-color: #1B7863; }
	#pr-summary { grid-template-columns: repeat(auto-fit, minmax(190px, 1fr)); }
	.pr-mini { display: flex; flex-wrap: wrap; gap: 4px; margin-top: 6px; }
	.pr-bar { height: 6px; margin-top: 8px; background: #e6ecea; border-radius: 999px; overflow: hidden; }
	.pr-bar span { display: block; height: 100%; background: #1B7863; border-radius: 999px; }

	.pr-panel { background: #fff; border: 1px solid #e1e7e5; border-radius: 14px; padding: 12px 14px 4px; box-shadow: 0 1px 3px rgba(23, 38, 42, .06); }
	.pr-panel .fbt-tools { margin-bottom: 10px; }
	.pr-scroll { overflow: auto; max-height: 70vh; border: 1px solid #e1e7e5; border-radius: 10px; }
	.pr-scroll #product-table { margin: 0; border-collapse: separate; border-spacing: 0; min-width: 100%; }
	/* Baris judul di dalam thead hanya untuk Export Excel */
	#product-table tr.pr-xls, #product-table tr.fbt-toolbar-row { display: none; }
	#product-table thead tr:last-child th { position: sticky; top: 0; z-index: 2; white-space: nowrap; border-width: 0 1px 0 0; border-color: #2f8a75; }
	#product-table td { border-width: 0 0 1px 0; border-color: #e6ecea; white-space: nowrap; vertical-align: middle; }
	#product-table tfoot td { position: sticky; bottom: 0; z-index: 1; background: #f1f7f5; font-weight: 700; border-top: 2px solid #1B7863; border-bottom: 0; }
	#product-table td.pr-num { text-align: right; }
	#product-table td.pr-code { font-weight: 600; color: #12584a; }
	#product-table td.pr-muted { color: #5d6d69; }
	#product-table td.pr-wrap { white-space: normal; min-width: 200px; }
	#product-table td.pr-act { width: 34px; padding: 6px 4px; text-align: center; }
	#product-table td.pr-act a { display: inline-flex; align-items: center; justify-content: center; width: 28px; height: 28px; border-radius: 6px; text-decoration: none; }
	#product-table td.pr-act a:hover { background: #e8f3f0; }
	#product-table td.pr-act img { display: block; }
	.pr-cancel { color: #b42318 !important; font-size: 16px; font-weight: 700; line-height: 1; border: 1px solid #f3c7c3; }
	#product-table td.pr-act a.pr-cancel:hover { background: #fef3f2; }
	.pr-dash { color: #b5c0bd; }
	.pr-badge { display: inline-block; padding: 2px 9px; border-radius: 999px; font-size: 11.5px; font-weight: 600; line-height: 1.5; background: #eef1f0; color: #4a5a56; }
	.pr-badge.ok { background: #e2f4ec; color: #136a4b; }
	.pr-badge.warn { background: #fff4d6; color: #8a5a00; }
	.pr-badge.bad { background: #fde8e6; color: #a8322a; }
	.pr-badge.info { background: #e6f0fb; color: #1f5a96; }
	.pr-empty { padding: 28px !important; text-align: center; color: #5d6d69; white-space: normal !important; }
	@media (max-width: 640px) {
		#wrapper { padding: 12px 16px; }
		.pr-title { font-size: 20px; }
		#pr-summary { grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 8px; }
		#pr-summary .fbt-card { padding: 10px 12px; }
		#pr-summary .fbt-card-value { font-size: 17px; overflow-wrap: anywhere; }
		.pr-panel { padding: 10px 10px 2px; }
	}
	@media print {
		body { background: #fff; }
		.pr-actions { display: none !important; }
		.pr-head, .pr-panel { border: 0; box-shadow: none; padding: 0; }
		.pr-scroll { overflow: visible; max-height: none; border: 0; }
		#product-table thead tr:last-child th, #product-table tfoot td { position: static; }
		.pr-badge { background: none !important; padding: 0; }
		#product-table td.pr-act, #product-table th.pr-act { display: none; }
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
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

'---- Escape tanda kutip untuk literal string SQL
Function SqlStr(v)
	SqlStr = Replace(Trim(v & ""), "'", "''")
End Function

'---- Encode output HTML (aman untuk Null)
Function H(v)
	If IsNull(v) Then H = "" Else H = Server.HTMLEncode(v & "")
End Function

'---- Encode nilai parameter URL (aman untuk Null)
Function U(v)
	If IsNull(v) Then U = "" Else U = Server.URLEncode(v & "")
End Function

'---- Tanggal MM/DD/YYYY yang valid, selain itu kosong (tanggal salah dulu membuat FormatDateTime error)
Function CleanMDY(v)
	Dim p, d
	CleanMDY = ""
	p = Split(Trim(v & ""), "/")
	If UBound(p) = 2 Then
		If IsNumeric(p(0)) And IsNumeric(p(1)) And IsNumeric(p(2)) Then
			If CInt(p(0)) >= 1 And CInt(p(0)) <= 12 Then
				d = DateSerial(CInt(p(2)), CInt(p(0)), CInt(p(1)))
				If Day(d) = CInt(p(1)) Then CleanMDY = Trim(v & "")
			End If
		End If
	End If
End Function

'---- Kelas warna badge status PR (hanya tampilan)
Function PrStateClass(v)
	Dim s
	s = LCase(Trim(v & ""))
	PrStateClass = ""
	If InStr(s, "cancel") > 0 Or InStr(s, "reject") > 0 Then
		PrStateClass = "bad"
	ElseIf InStr(s, "approv") > 0 Or InStr(s, "order") > 0 Or InStr(s, "close") > 0 Or InStr(s, "complete") > 0 Or InStr(s, "release") > 0 Then
		PrStateClass = "ok"
	ElseIf InStr(s, "progress") > 0 Or InStr(s, "authoriz") > 0 Or InStr(s, "confirm") > 0 Then
		PrStateClass = "info"
	ElseIf InStr(s, "plan") > 0 Or InStr(s, "draft") > 0 Or InStr(s, "new") > 0 Then
		PrStateClass = "warn"
	End If
End Function

'---- Chip filter di kop laporan
Function Chip(lbl, v)
	Chip = ""
	If Len(Trim(v & "")) > 0 Then Chip = "<span class=""pr-chip""><b>" & lbl & "</b>" & H(v) & "</span>"
End Function

'---- Parameter filter: tanda kutip di-escape sebelum disambung ke SQL
v_txtPRDATE_From=CleanMDY(request("PRDATE_From"))
v_txtPRDATE_To=CleanMDY(request("PRDATE_To"))
v_txtPRCODE=SqlStr(request("PRCODE"))
v_txtITEMCODE=SqlStr(request("ITEMCODE"))
v_txtRequestorID=SqlStr(request("REQUESTORID"))
v_txtRowStateId=SqlStr(request("ROWSTATEID"))

businessunit= SqlStr(request("businessunit"))
'---- Perbaikan: nama PT kosong bila laporan dibuka tanpa parameter businessunit
if len(businessunit) = 0 then businessunit = SqlStr(request("fcba"))
if len(businessunit) = 0 then businessunit = SqlStr(Session("fcba"))

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
'---- Nama perusahaan kosong di master BusinessUnit: pakai nama perusahaan dari sesi login
if len(trim(CompanyName & "")) = 0 then CompanyName = Session("COMPANY")
if len(trim(fcbaname & "")) = 0 then fcbaname = Session("BA")
'---- Hindari "PT. PT. ..." bila nama sudah diawali PT
if UCase(Left(Trim(CompanyName & ""), 3)) = "PT." or UCase(Left(Trim(CompanyName & ""), 3)) = "PT " then CompanyName = Trim(Mid(Trim(CompanyName & ""), 4))

''		"       stock_api.get_stockName (b.itemcode, b.fcba) item_description," &  _'
txtSQL = 	"	SELECT DISTINCT fcba,PRDATE," &  _
			"       PRCODE," &  _
			"       REQUESTORID," &  _
			"       PURCHASEGROUPID," &  _
			"       MANAGEDID," &  _
			"       ROWSTATEID," &  _
			"       remarks," &  _
			"       lineid," &  _
			"       ITEMCODE," &  _
			"       fcname," &  _
			"       item_description," &  _
			"       requestedqty," &  _
			"       approvedqty," &  _
			"       orderedqty," &  _
			"       unitofmeasurement, " &  _
			"       nilai " &  _
			"  FROM (" & _
			"	SELECT a.fcba,a.PRDATE," &  _
			"       a.PRCODE," &  _
			"       a.REQUESTORID," &  _
			"       a.PURCHASEGROUPID," &  _
			"       a.MANAGEDID," &  _
			"       a.ROWSTATEID," &  _
			"       a.remarks," &  _
			"       b.lineid," &  _
			"       b.itemcode," &  _
			"       dc.fcname," &  _
			"		s.FCNAME item_description, " &  _
			"       NVL (b.requestedqty, 0) requestedqty," &  _
			"       NVL (b.approvedqty, 0) approvedqty," &  _
			"       NVL (b.orderedqty, 0) orderedqty," &  _
			"       b.unitofmeasurement, " &  _
			"       (SELECT count(*) FROM PURCHASEREQUESTDETAILS_HIST WHERE CURRENTAUTHORIZELEVEL = '900' AND prcode = a.PRCODE AND FCENTRY IN ('DEPT_HEAD_PURCH', 'RINTO_PURNAMASIDI')) as nilai" &  _
			"  FROM purchaserequest a" &  _
			"       INNER JOIN purchaserequestdetails b" &  _
			" INNER JOIN stock s on (s.FCCODE=b.ITEMCODE and s.fcba=b.fcba) " & _
			"          ON (a.prcode = b.prcode AND a.fcba = b.fcba)" &  _
			"       LEFT JOIN document_classification dc" &  _
			"          ON (dc.fccode = a.document_classification)" &  _
			" WHERE     a.prcode =  b.prcode AND a.rowstateid not in ('Cancelled', 'Rejected') " &  _
			"       AND a.fcba = '" &businessunit & "' AND dc.DOCUMENT_TYPE='PR'" 'added by sendy 10/29/2018
 
if  len(trim(v_txtPRDATE_From)) > 0 then
    txtSQL = txtSQL & " and a.prdate >=to_Date('" & v_txtPRDATE_From & "','MM/DD/YYYY')"
	txtSelect = txtSelect & " From " & FormatDateTime(v_txtPRDATE_From,1)	
end if
if  len(trim(v_txtPRDATE_To)) > 0 then
    txtSQL =txtSQL &  " and a.prdate <=to_Date('" & v_txtPRDATE_To & "','MM/DD/YYYY')"
	txtSelect = txtSelect & " To " & FormatDateTime(v_txtPRDATE_To,1)	
end if
if  len(trim(v_txtPRCODE)) > 0 then
    txtSQL = txtSQL & " and a.prcode='" & v_txtPRCODE & "'"
	txtSelect = txtSelect & " PR Code " & v_txtPRCODE	
end if
if  len(trim(v_txtITEMCODE)) > 0 then
    txtSQL =txtSQL &  " and b.itemcode='" & v_txtITEMCODE & "'"
	txtSelect = txtSelect & " Item Code " & v_txtITEMCODE	
end if 
if  len(trim(v_txtRequestorID)) > 0 then
    txtSQL =txtSQL &  " and a.RequestorId='" & v_txtRequestorID & "'"
	txtSelect = txtSelect & " Requestor Code " & v_txtRequestorID
end if 

if  len(trim(v_txtRowStateId)) > 0 then
    txtSQL =txtSQL &  " and a.ROWSTATEID='" & v_txtRowStateId & "'"
	txtSelect = txtSelect & " Status  " & v_txtRowStateId	
end if

	txtSQL = txtSQL & " Order by b.prcode ,to_number(b.lineid))" 'ditambah ukie tgl 9 okt 2013 untuk Order by
	
	txtSQL = txtSQL & " Order by prcode ,lineid"

if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
end if

 'response.Write(txtsql)
response.Flush()
%>
	<body>
		<div id="loading-overlay" style="display:none; position:fixed; top:0; left:0; width:100%; height:100%;
			background:rgba(0,0,0,0.5); z-index:9999; text-align:center;">
			<div style="position:absolute; top:50%; left:50%; transform:translate(-50%,-50%);
						background:#fff; padding:30px 40px; border-radius:8px; box-shadow:0 2px 10px rgba(0,0,0,0.3);">
				<div id="loadingjudul" style="font-weight:bold; font-size:14px; margin-bottom:10px; color:#333;">
					Memproses...
				</div>
				<div id="loading">
					<img src="IMAGES/loading.gif" alt="Loading..." onerror="this.style.display='none'">
					<span style="font-size:12px; color:#666;">Mohon tunggu sebentar</span>
				</div>
			</div>
		</div>
		<div id="wrapper">
		<%
		Set rsCat = Server.CreateObject("ADODB.Recordset")
		rsCat.CursorType = 3
		rsCat.CursorLocation = 3
		rsCat.Open txtSQL, connstr

		if Not rsCat.eof then
			nCountofField=rsCat.fields.count-1
		end if
		NO=1
		%>
		<div class="pr-head">
			<div>
				<div class="pr-company">PT. <%=H(CompanyName)%></div>
				<div class="pr-ba"><%=H(fcbaname)%></div>
				<h2 class="pr-title">Purchase Request</h2>
				<div class="pr-chips">
					<%if len(v_txtPRDATE_From) > 0 or len(v_txtPRDATE_To) > 0 then%>
					<span class="pr-chip"><b>Periode</b><%=H(v_txtPRDATE_From)%> &ndash; <%=H(v_txtPRDATE_To)%></span>
					<%else%>
					<span class="pr-chip"><b>Periode</b>Semua tanggal</span>
					<%end if%>
					<%=Chip("PR Code", request("PRCODE"))%>
					<%=Chip("Item", request("ITEMCODE"))%>
					<%=Chip("Requestor", request("REQUESTORID"))%>
					<%=Chip("Status", request("ROWSTATEID"))%>
				</div>
			</div>
			<div class="pr-actions">
				<button type="button" class="pr-print" onclick="window.print()">Cetak</button>
				<!--#include file="print_excel.asp"-->
			</div>
		</div>
		<div class="fbt-summary" id="pr-summary"></div>
		<div class="pr-panel">
		<div id="pr-tools"></div>
		<div class="pr-scroll">
		<table border="0" width="100%" id="product-table">
			<thead>
			<tr class="pr-xls">
				<th colspan="16" style=" background-color:#FFFFFF; color:#000000; border:none">
					<h4>PT. <%=H(CompanyName)%></h4>
					<h5><%=H(fcbaname)%></h5>
					<h2>PURCHASE REQUEST</h2>
				</th>
			</tr>
			<tr class="pr-xls">
				<th colspan="16" style=" background-color:#FFFFFF; color:#000000; border:none"><%=H(txtSelect)%></th>
			</tr>
			<tr>
				<th colspan="3" class="pr-act" data-nosort>Aksi</th>
				<th>Date</th>
				<th>PR.Code</th>
				<th>Requestor</th>
				<th>Status</th>
				<th>Remarks</th>
				<th>Line Id</th>
				<th>Stock Code</th>
				<th>Stock Description</th>
				<th>Uom</th>
				<th>Classification</th>
				<th>Requested (Qty)</th>
				<th>Approved (Qty)</th>
				<th>Ordered (Qty)</th>
			</tr>
			</thead>
			<tbody>
				<%
				Dim v_Opening,v_Debit,v_Credit,v_Saldo,v_Quantity
				v_Quantity=0
				v_Amount = 0
				v_req = 0
				v_appr = 0
				v_ord = 0
				While Not rsCat.EOF%>
					<tr>
						<td class="pr-act">
							<a title="Edit PR" href="PurchaseRequestDetails.Asp?PRCODE=<%=U(rscat("prcode"))%>&fcba=<%=U(rscat("fcba"))%>&PrDate=<%=U(rscat("prdate"))%>&RowStateId=<%=U(rscat("RowstateId"))%>&txtJenis=Edit">
								<img src="IMAGES/edit.gif" alt="Edit Data : <%=H(rscat(1))%>" border="0">
							</a>
						</td>
						<td class="pr-act">
							<a title="Cetak PR" href="javascript:launchwin('PrintPR.Asp?PrCode=<%=U(rscat("prcode"))%>&fcba=<%=U(businessunit)%>&PrDate=<%=U(rscat("prdate"))%>&Decription=<%=U(Description)%>&RequestorId=<%=U(rscat("RequestorId"))%>&PurchaseGroupId=<%=U(rscat("PURCHASEGROUPID"))%>&ManagedId=<%=U(rscat("MANAGEDID"))%>&Remarks=<%=U(rscat("item_description"))%>&RowStateId=<%=U(rscat("RowstateId"))%>&Mode=Print','newwindow','width='+(980)+',height='+(650)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">
								<img src="IMAGES/print.gif" alt="Print Data : <%=H(rscat(1))%>" border="0">
							</a>
						</td>
						<td class="pr-act">
							<%If rscat("RowstateId") = "Planned" or rscat("nilai") = "0" then%>
								<a href="#" class="pr-cancel" title="Batalkan PR <%=H(rscat("prcode"))%>" data-pr="<%=H(rscat("prcode"))%>" onclick="cancelPR(this.getAttribute('data-pr')); return false;">&times;</a>
							<%else%>
								<span class="pr-dash">-</span>
							<%end if%>
						</td>
						<td><%=H(rscat("prdate"))%></td>
						<td class="pr-code"><%=H(rscat("prcode"))%></td>
						<td><%=H(rscat("RequestorId"))%></td>
						<td><%if len(rscat("RowstateId") & "") > 0 then%><span class="pr-badge <%=PrStateClass(rscat("RowstateId"))%>"><%=H(rscat("RowstateId"))%></span><%end if%></td>
						<td class="pr-wrap"><%=H(rscat("Remarks"))%></td>
						<td class="pr-num pr-muted"><%=H(rscat("Lineid"))%></td>
						<td class="pr-muted"><%=H(rscat("ItemCode"))%></td>
						<td class="pr-wrap"><%=H(rscat("item_description"))%></td>
						<td><%=H(rscat("unitofmeasurement"))%></td>
						<td><%=H(rscat("fcname"))%></td>
						<td class="pr-num"><%=FormatNumber(cdbl(rscat("requestedqty")),2)%></td>
						<td class="pr-num"><%=FormatNumber(cdbl(rscat("approvedqty")),2)%></td>
						<td class="pr-num"><%=FormatNumber(cdbl(rscat("orderedqty")),2)%></td>
					</tr>
					<%
					v_req = v_req + cdbl(rscat("requestedqty"))
					v_appr = v_appr + cdbl(rscat("approvedqty"))
					v_ord = v_ord + cdbl(rscat("orderedqty"))
					rsCat.MoveNext
					'response.Flush()
					no=no+1
				Wend
				rsCat.close()
				Set rsCat = Nothing
				if NO = 1 then%>
				<tr><td colspan="16" class="pr-empty">Tidak ada Purchase Request untuk filter ini.</td></tr>
				<%end if%>
			</tbody>
			<tfoot>
			<tr bgcolor="#F4F4A0" style="font-weight: bold;">
				<td colspan="13">Total (<span class="fbt-count"><%=NO-1%></span> baris)</td>
				<td class="pr-num" data-sum="13"><%=FormatNumber(v_req,2)%></td>
				<td class="pr-num" data-sum="14"><%=FormatNumber(v_appr,2)%></td>
				<td class="pr-num" data-sum="15"><%=FormatNumber(v_ord,2)%></td>
			</tr>
			</tfoot>
		</table>
		</div>
		<div id="pr-foot"></div>
		</div>
		<input  type="hidden" value="<%=H(txtTable)%>" name="txtTable">
		<input  type="hidden" value="<%=H(txtModule)%>" name="txtModule">
		</div>
	<script src="/js/jquery-1.9.1.js"></script>
	<script src="/js/jquery-ui.js"></script>
	<link rel="stylesheet" href="css/jquery-ui.css" />
	<link href="/js/select2/select2.min.css" rel="stylesheet" type="text/css"/>
	<script src="/js/select2/select2.min.js" type="text/javascript"></script>
	<script src="https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js"></script>
	<script type="text/javascript">
		function cancelPR(prcode) {
			if (confirm("Are you sure you want to cancel PR: " + prcode + "?")) {
				$('#loadingjudul').text('Membatalkan PR: ' + prcode + '...');
				$('#loading-overlay').show();

				$.post("updatedata.asp?idview=1" + "&prcode=" + encodeURIComponent(prcode), { prcode: prcode }, function(data) {
					$('#loading-overlay').hide();
					try {
						var json = JSON.parse(data);
						if (json.status === "success") {
							alert(json.message);
							location.reload(); // refresh tabel setelah cancel
						} else {
							alert("Error: " + json.message);
						}
					} catch(e) {
						alert("Response tidak valid dari server.");
					}
				}).fail(function() {
					$('#loading-overlay').hide();
					alert("Request gagal, coba lagi.");
				});
			}
		}

		// Tabel ala DataTables + kartu ringkasan (hanya tampilan; angka dari baris yang dirender server)
		window.addEventListener('load', function () {
			var table = document.getElementById('product-table');
			var summary = document.getElementById('pr-summary');
			if (!window.FbpTable || !table) return;   // tetap dipasang walau data kosong: toolbar cari & jumlah baris tetap tampil
			function fmt(n, d) { return n.toFixed(d || 0).replace(/\B(?=(\d{3})+(?!\d))/g, ','); }
			function txt(r, i) { return (r.cells[i].textContent || '').replace(/\s+/g, ' ').trim(); }
			function num(r, i) { return parseFloat(txt(r, i).replace(/,/g, '')) || 0; }
			function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
			function card(label, value, note, extra) {
				return '<div class="fbt-card' + (extra || '') + '"><div class="fbt-card-label">' + label + '</div>' +
					'<div class="fbt-card-value">' + value + '</div><div class="fbt-card-note">' + note + '</div></div>';
			}
			FbpTable.patchExport();
			FbpTable.init(table, {
				pageSize: 20,
				onChange: function (rows) {
					var s = { 13: 0, 14: 0, 15: 0 }, pr = {}, nPr = 0, st = {}, req = {}, nReq = 0, open = 0;
					rows.forEach(function (r) {
						s[13] += num(r, 13); s[14] += num(r, 14); s[15] += num(r, 15);
						if (num(r, 14) > num(r, 15)) open++;
						var k = txt(r, 4);
						if (!pr[k]) { pr[k] = txt(r, 6) || '-'; nPr++; st[pr[k]] = (st[pr[k]] || 0) + 1; }
						var q = txt(r, 5);
						if (q && !req[q]) { req[q] = 1; nReq++; }
					});
					Array.prototype.forEach.call(table.tFoot.querySelectorAll('[data-sum]'), function (td) {
						td.textContent = fmt(s[td.getAttribute('data-sum')], 2);
					});
					var pct = s[14] ? Math.min(100, Math.round(s[15] * 100 / s[14])) : 0;
					var badges = Object.keys(st).sort(function (a, b) { return st[b] - st[a]; }).slice(0, 4).map(function (k) {
						var c = /cancel|reject/i.test(k) ? 'bad' : /approv|order|close|complete|release/i.test(k) ? 'ok' : /progress|authoriz|confirm/i.test(k) ? 'info' : /plan|draft|new/i.test(k) ? 'warn' : '';
						return '<span class="pr-badge ' + c + '">' + esc(k) + ' ' + st[k] + '</span>';
					}).join('');
					summary.innerHTML =
						card('Purchase Request', fmt(nPr), rows.length + ' baris item &middot; ' + nReq + ' requestor') +
						card('Item belum di-PO', fmt(open), 'qty disetujui &gt; qty sudah di-order', open ? ' accent' : '') +
						card('Progres order', pct + '%', 'qty di-order dari qty disetujui<div class="pr-bar"><span style="width:' + pct + '%"></span></div>') +
						card('Status PR', nPr ? fmt(nPr) + ' PR' : '-', '<div class="pr-mini">' + badges + '</div>');
				}
			});
			// Toolbar & pager di luar area gulir agar tetap terlihat walau tabel lebar
			var tools = table.querySelector('.fbt-toolbar-row .fbt-tools'), foot = table.parentNode.querySelector('.fbt-foot');
			if (tools) document.getElementById('pr-tools').appendChild(tools);
			if (foot) document.getElementById('pr-foot').appendChild(foot);
		});
	</script>
		<!--#include file="Bottom_Report.asp"-->
