<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>FBPlants Reporting - BKM UnApproval</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<link rel="stylesheet" href="/css/fbp-table.css?v=20261010c" type="text/css" />
<script type="text/javascript" src="/js/fbp-table.js?v=20261009c"></script>
<style>
	/* Laporan BKM UnApproval (Replanned): kop, kartu ringkasan, tabel ala DataTables dengan header lengket */
	body { background: #f3f6f5; }
	.br-head {
		display: flex; flex-wrap: wrap; align-items: flex-start; justify-content: space-between; gap: 16px;
		padding: 18px 22px; background: #fff; border: 1px solid #e1e7e5; border-radius: 14px;
		border-left: 5px solid #1B7863; box-shadow: 0 1px 3px rgba(23, 38, 42, .06);
	}
	.br-company { font-size: 12px; font-weight: 700; color: #5d6d69; text-transform: uppercase; letter-spacing: .6px; }
	.br-ba { margin-top: 2px; font-size: 12px; color: #8a9a96; }
	.br-title { margin: 6px 0 0; font-size: 24px; font-weight: 700; color: #12584a; letter-spacing: .2px; }
	.br-sub { margin-top: 2px; font-size: 13px; color: #5d6d69; }
	.br-chips { display: flex; flex-wrap: wrap; gap: 6px; margin-top: 10px; }
	.br-chip { display: inline-flex; align-items: center; gap: 6px; padding: 3px 10px; font-size: 12px; color: #17262a; background: #eef5f3; border: 1px solid #d3e6e0; border-radius: 999px; }
	.br-chip b { font-weight: 600; color: #5d6d69; }
	.br-actions { display: flex; align-items: center; gap: 8px; }
	.br-actions input[type="button"] { float: none !important; margin: 0 !important; }
	.br-print { min-height: 36px; padding: 0 14px; font: inherit; font-size: 13px; font-weight: 600; color: #1B7863; background: #fff; border: 1px solid #b9d9d0; border-radius: 8px; cursor: pointer; }
	.br-print:hover, .br-print:focus { background: #e8f3f0; border-color: #1B7863; }
	#br-summary { grid-template-columns: repeat(auto-fit, minmax(190px, 1fr)); }
	.br-mini { display: flex; flex-wrap: wrap; gap: 4px; margin-top: 6px; }

	.br-panel { background: #fff; border: 1px solid #e1e7e5; border-radius: 14px; padding: 12px 14px 4px; box-shadow: 0 1px 3px rgba(23, 38, 42, .06); }
	.br-panel .fbt-tools { margin-bottom: 10px; }
	.br-scroll { overflow: auto; max-height: 70vh; border: 1px solid #e1e7e5; border-radius: 10px; }
	.br-scroll #product-table { margin: 0; border-collapse: separate; border-spacing: 0; min-width: 100%; }
	/* Baris judul di dalam thead hanya untuk Export Excel */
	#product-table tr.br-xls, #product-table tr.fbt-toolbar-row { display: none; }
	#product-table thead tr:last-child th { position: sticky; top: 0; z-index: 2; white-space: nowrap; border-width: 0 1px 0 0; border-color: #2f8a75; }
	#product-table td { border-width: 0 0 1px 0; border-color: #e6ecea; white-space: nowrap; vertical-align: middle; }
	#product-table tfoot td { position: sticky; bottom: 0; z-index: 1; background: #f1f7f5; font-weight: 700; border-top: 2px solid #1B7863; border-bottom: 0; }
	#product-table td.br-code { font-weight: 600; color: #12584a; }
	#product-table td.br-muted { color: #5d6d69; }
	.br-badge { display: inline-block; padding: 2px 9px; border-radius: 999px; font-size: 11.5px; font-weight: 600; line-height: 1.5; background: #eef1f0; color: #4a5a56; }
	.br-badge.ok { background: #e2f4ec; color: #136a4b; }
	.br-badge.warn { background: #fff4d6; color: #8a5a00; }
	.br-badge.info { background: #e6f0fb; color: #1f5a96; }
	.br-badge.bad { background: #fde8e6; color: #a8322a; }
	.br-empty { padding: 28px !important; text-align: center; color: #5d6d69; white-space: normal !important; }
	@media (max-width: 640px) {
		#wrapper { padding: 12px 16px; }
		.br-title { font-size: 20px; }
		#br-summary { grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 8px; }
		#br-summary .fbt-card { padding: 10px 12px; }
		#br-summary .fbt-card-value { font-size: 17px; overflow-wrap: anywhere; }
		.br-panel { padding: 10px 10px 2px; }
	}
	@media print {
		body { background: #fff; }
		.br-actions { display: none !important; }
		.br-head, .br-panel { border: 0; box-shadow: none; padding: 0; }
		.br-scroll { overflow: visible; max-height: none; border: 0; }
		#product-table thead tr:last-child th, #product-table tfoot td { position: static; }
		.br-badge { background: none !important; padding: 0; }
	}
</style>
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

'---- Tanggal MM/DD/YYYY yang valid, selain itu kosong (tanggal salah dulu membuat to_date error)
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

'---- Nama jenis BKM dari document_classification
Function BkmTypeName(v)
	Select Case Trim(v & "")
		Case "501" : BkmTypeName = "BKM Perawatan"
		Case "502" : BkmTypeName = "BKM Traksi"
		Case "503" : BkmTypeName = "BKM Panen"
		Case "504" : BkmTypeName = "BKM WIP"
		Case Else : BkmTypeName = Trim(v & "")
	End Select
End Function

'---- Kelas warna badge status BKM (hanya tampilan)
Function BkmStateClass(v)
	Dim s
	s = LCase(Trim(v & ""))
	BkmStateClass = ""
	If InStr(s, "approv") > 0 Or InStr(s, "post") > 0 Then
		BkmStateClass = "ok"
	ElseIf InStr(s, "review") > 0 Then
		BkmStateClass = "info"
	ElseIf InStr(s, "plan") > 0 Then
		BkmStateClass = "warn"
	ElseIf InStr(s, "cancel") > 0 Or InStr(s, "reject") > 0 Then
		BkmStateClass = "bad"
	End If
End Function

'---- Chip filter di kop laporan
Function Chip(lbl, v)
	Chip = ""
	If Len(Trim(v & "")) > 0 Then Chip = "<span class=""br-chip""><b>" & lbl & "</b>" & H(v) & "</span>"
End Function

businessunit=SqlStr(request("businessunit"))

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

 dim bkmfrom,bkmto,txtfcba1,txtafd,txtstatus1,filter,document_clasification
	'---- Parameter filter: tanggal divalidasi, teks di-escape sebelum disambung ke SQL
    bkmfrom     = CleanMDY(Request("entry_From"))
    bkmto       = CleanMDY(Request("entry_To"))
    txtfcba1    = SqlStr(Request("businessunit"))
    txtafd      = SqlStr(Request("sectionname"))
    txtstatus1  = SqlStr(Request("rowstate"))
    modulname  = SqlStr(Request("modulname"))
    usercode   = SqlStr(Session("CodeName"))

    if trim(txtafd) = "" then
		filter = " AND hdr.sectioncode in(select SECTIONNAME  from user_section_tab where FCUSERCODE='" & usercode & "' ) "
	else
	    filter = " AND hdr.sectioncode in(select SECTIONNAME  from user_section_tab where FCUSERCODE='" & usercode & "' AND sectionname='" & txtafd & "') "
	end if

	q = " SELECT hdr.fcba, hdr.bkmno,hdr.rkbno, hdr.bkmdate, hdr.document_classification, hdr.sectioncode, hdr.gangcode,  hdr.document_classification," & _
		" hdr.VEHICLECODE,hdr.fcentry,hdr.fcedit,aus.authorizelevelcode,hdr.rowstate, " & _
		" hdr.fcedit,hdr.lastupdate" & _
	    " FROM bkmhdr hdr" & _
	    " LEFT JOIN authorizetemplate at ON (at.document_type = 'AA' AND at.document_classification = hdr.document_classification)" & _
		" INNER JOIN authorizestructure aus ON (aus.fcba = hdr.fcba AND aus.authorizetemplatecode = at.fccode AND aus.fcusercode = '" & usercode & "')" & _
	    " WHERE hdr.bkmdate between to_date('" & bkmfrom & "','mm/dd/yyyy') AND to_date('" & bkmto & "','mm/dd/yyyy') AND hdr.rowstate = 'Planned'" & _
		" AND at.document_classification = hdr.document_classification AND hdr.fcba='" & txtfcba1 & "' AND hdr.STATUSREPLANNED='0' AND hdr.document_classification='" & modulname & "' " & _
        filter & _
		" ORDER BY hdr.bkmdate, hdr.fcba, hdr.bkmno"

	if len(bkmfrom) > 0 or len(bkmto) > 0 then txtSelect = " From " & bkmfrom & " To " & bkmto
	if len(trim(txtafd)) > 0 then txtSelect = txtSelect & " Section " & Request("sectionname")
	if len(trim(txtSelect)) > 0 then
	txtSelect = "Selected  by : " & txtSelect
	else
	txtSelect = "Selected by All Data"
	end if

'' response.Write(q)
response.Flush()
%>
	<body>
		<div id="wrapper">
		<%
		Set rsCat = Server.CreateObject("ADODB.Recordset")
		rsCat.CursorType = 3
		rsCat.CursorLocation = 3
		rsCat.Open q, connstr

		if Not rsCat.eof then
		   nCountofField=rsCat.fields.count-1
		end if
		NO=1
		%>
		<div class="br-head">
			<div>
				<div class="br-company">PT. <%=H(CompanyName)%></div>
				<div class="br-ba"><%=H(fcbaname)%></div>
				<h2 class="br-title">BKM UnApproval</h2>
				<div class="br-sub">BKM berstatus Planned yang dikembalikan (un-approve) dan menunggu proses replan</div>
				<div class="br-chips">
					<%if len(bkmfrom) > 0 or len(bkmto) > 0 then%>
					<span class="br-chip"><b>Periode</b><%=H(bkmfrom)%> &ndash; <%=H(bkmto)%></span>
					<%end if%>
					<%=Chip("Jenis BKM", BkmTypeName(Request("modulname")))%>
					<%=Chip("Section", Request("sectionname"))%>
				</div>
			</div>
			<div class="br-actions">
				<button type="button" class="br-print" onclick="window.print()">Cetak</button>
				<!--#include file="print_excel.asp"-->
			</div>
		</div>
		<div class="fbt-summary" id="br-summary"></div>
		<div class="br-panel">
		<div id="br-tools"></div>
		<div class="br-scroll">
		<table border="0" width="100%" id="product-table">
			<thead>
			<tr class="br-xls">
				<th colspan="11" style=" background-color:#FFFFFF; color:#000000; border:none">
					<h4>PT. <%=H(CompanyName)%></h4>
					<h5><%=H(fcbaname)%></h5>
					<h2>BKM UN APPROVAL</h2>
				</th>
			</tr>
			<tr class="br-xls">
				<th colspan="11" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=H(txtSelect)%></th>
			</tr>
			<tr>
				<th>Kebun/Pabrik</th>
				<th>BKM Date</th>
				<th>BKM No</th>
				<th>BKM RKB</th>
				<th>BKM Type</th>
				<th>Section Name</th>
				<th>Gang</th>
				<th>Vehicle Code</th>
				<th>Status</th>
				<th>User UnApproved</th>
				<th>Last Time UnApproved</th>
			</tr>
			</thead>
			<tbody>
				<%
				Dim bkmtype
				While Not rsCat.EOF
					bkmtype = BkmTypeName(rscat("document_classification"))
				%>
				<tr>
					<td class="br-muted"><%=H(rscat("fcba"))%></td>
					<td><%=H(rscat("bkmdate"))%></td>
					<td class="br-code"><%=H(rscat("bkmno"))%></td>
					<td class="br-muted"><%=H(rscat("rkbno"))%></td>
					<td><%=H(bkmtype)%></td>
					<td><%=H(rscat("sectioncode"))%></td>
					<td><%=H(rscat("gangcode"))%></td>
					<td><%=H(rscat("vehiclecode"))%></td>
					<td><%if len(rscat("rowstate") & "") > 0 then%><span class="br-badge <%=BkmStateClass(rscat("rowstate"))%>"><%=H(rscat("rowstate"))%></span><%end if%></td>
					<td><%=H(rscat("fcedit"))%></td>
					<td><%=H(rscat("lastupdate"))%></td>
				</tr>
				<%
					rsCat.MoveNext
					'response.Flush()
					no=no+1
				Wend
				rsCat.close()
				Set rsCat = Nothing
				if NO = 1 then%>
				<tr><td colspan="11" class="br-empty">Tidak ada BKM UnApproval untuk filter ini.</td></tr>
				<%end if%>
			</tbody>
			<tfoot>
			<tr bgcolor="#F4F4A0" style="font-weight: bold;">
				<td colspan="11">Total (<span class="fbt-count"><%=NO-1%></span> BKM)</td>
			</tr>
			</tfoot>
		</table>
		</div>
		<div id="br-foot"></div>
		</div>
		<input  type="hidden" value="<%=H(txtTable)%>" name="txtTable">
		<input  type="hidden" value="<%=H(txtModule)%>" name="txtModule">
		</div>
	<script type="text/javascript">
	// Tabel ala DataTables + kartu ringkasan (hanya tampilan; data dari baris yang dirender server)
	window.addEventListener('load', function () {
		var table = document.getElementById('product-table');
		var summary = document.getElementById('br-summary');
		if (!window.FbpTable || !table) return;   // tetap dipasang walau data kosong: toolbar cari & jumlah baris tetap tampil
		function txt(r, i) { return (r.cells[i].textContent || '').replace(/\s+/g, ' ').trim(); }
		function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
		function count(rows, i) { var m = {}; rows.forEach(function (r) { var k = txt(r, i) || '-'; m[k] = (m[k] || 0) + 1; }); return m; }
		function top(m, n, cls) {
			return Object.keys(m).sort(function (a, b) { return m[b] - m[a]; }).slice(0, n).map(function (k) {
				return '<span class="br-badge ' + (cls || '') + '">' + esc(k) + ' ' + m[k] + '</span>';
			}).join('');
		}
		function card(label, value, note, extra) {
			return '<div class="fbt-card' + (extra || '') + '"><div class="fbt-card-label">' + label + '</div>' +
				'<div class="fbt-card-value">' + value + '</div><div class="fbt-card-note">' + note + '</div></div>';
		}
		FbpTable.patchExport();
		FbpTable.init(table, {
			pageSize: 20,
			onChange: function (rows) {
				var sec = count(rows, 5), gang = count(rows, 6), usr = count(rows, 9), day = count(rows, 1);
				var nSec = Object.keys(sec).length, nGang = Object.keys(gang).length;
				summary.innerHTML =
					card('BKM UnApproval', rows.length, Object.keys(day).length + ' tanggal BKM', rows.length ? ' accent' : '') +
					card('Section', rows.length ? nSec : 0, '<div class="br-mini">' + top(sec, 4, 'warn') + '</div>') +
					card('Gang', rows.length ? nGang : 0, rows.length ? 'gang yang BKM-nya dikembalikan' : '-') +
					card('User UnApproved', rows.length ? Object.keys(usr).length : 0, '<div class="br-mini">' + top(usr, 3, 'info') + '</div>');
			}
		});
		// Toolbar & pager di luar area gulir agar tetap terlihat walau tabel lebar
		var tools = table.querySelector('.fbt-toolbar-row .fbt-tools'), foot = table.parentNode.querySelector('.fbt-foot');
		if (tools) document.getElementById('br-tools').appendChild(tools);
		if (foot) document.getElementById('br-foot').appendChild(foot);
	});
	</script>
		<!--#include file="Bottom_Report.asp"-->
