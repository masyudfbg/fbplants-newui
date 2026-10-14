<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>FBPlants Reporting - Monitoring Stock</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<link rel="stylesheet" href="/css/fbp-table.css?v=20261010c" type="text/css" />
<script type="text/javascript" src="/js/fbp-table.js?v=20261009c"></script>
<style>
	/* Laporan Monitoring Stock: kop, kartu ringkasan, tabel ala DataTables dengan header lengket */
	body { background: #f3f6f5; }
	.ms-head {
		display: flex; flex-wrap: wrap; align-items: flex-start; justify-content: space-between; gap: 16px;
		padding: 18px 22px; background: #fff; border: 1px solid #e1e7e5; border-radius: 14px;
		border-left: 5px solid #1B7863; box-shadow: 0 1px 3px rgba(23, 38, 42, .06);
	}
	.ms-company { font-size: 12px; font-weight: 700; color: #5d6d69; text-transform: uppercase; letter-spacing: .6px; }
	.ms-ba { margin-top: 2px; font-size: 12px; color: #8a9a96; }
	.ms-title { margin: 6px 0 0; font-size: 24px; font-weight: 700; color: #12584a; letter-spacing: .2px; }
	.ms-chips { display: flex; flex-wrap: wrap; gap: 6px; margin-top: 10px; }
	.ms-chip { display: inline-flex; align-items: center; gap: 6px; padding: 3px 10px; font-size: 12px; color: #17262a; background: #eef5f3; border: 1px solid #d3e6e0; border-radius: 999px; }
	.ms-chip b { font-weight: 600; color: #5d6d69; }
	.ms-actions { display: flex; align-items: center; gap: 8px; }
	.ms-actions input[type="button"] { float: none !important; margin: 0 !important; }
	.ms-print { min-height: 36px; padding: 0 14px; font: inherit; font-size: 13px; font-weight: 600; color: #1B7863; background: #fff; border: 1px solid #b9d9d0; border-radius: 8px; cursor: pointer; }
	.ms-print:hover, .ms-print:focus { background: #e8f3f0; border-color: #1B7863; }
	#ms-summary { grid-template-columns: repeat(auto-fit, minmax(190px, 1fr)); }
	.fbt-card.warn { border-top-color: #d9822b; }

	.ms-panel { background: #fff; border: 1px solid #e1e7e5; border-radius: 14px; padding: 12px 14px 4px; box-shadow: 0 1px 3px rgba(23, 38, 42, .06); }
	.ms-panel .fbt-tools { margin-bottom: 10px; }
	.ms-scroll { overflow: auto; max-height: 70vh; border: 1px solid #e1e7e5; border-radius: 10px; }
	.ms-scroll #product-table { margin: 0; border-collapse: separate; border-spacing: 0; min-width: 100%; }
	/* Baris judul di dalam thead hanya untuk Export Excel */
	#product-table tr.ms-xls, #product-table tr.fbt-toolbar-row { display: none; }
	#product-table thead tr:last-child th { position: sticky; top: 0; z-index: 2; white-space: nowrap; border-width: 0 1px 0 0; border-color: #2f8a75; }
	#product-table td { border-width: 0 0 1px 0; border-color: #e6ecea; white-space: nowrap; vertical-align: middle; }
	#product-table tfoot td { position: sticky; bottom: 0; z-index: 1; background: #f1f7f5; font-weight: 700; border-top: 2px solid #1B7863; border-bottom: 0; }
	#product-table td.ms-num { text-align: right; }
	#product-table td.ms-end { text-align: right; font-weight: 700; }
	#product-table td.ms-code { font-weight: 600; color: #12584a; }
	#product-table td.ms-muted { color: #5d6d69; }
	#product-table td.ms-zero { color: #b5c0bd; }
	#product-table td.ms-wrap { white-space: normal; min-width: 220px; }
	#product-table td.ms-act { width: 34px; padding: 6px 4px; text-align: center; }
	#product-table td.ms-act a { display: inline-flex; align-items: center; justify-content: center; width: 28px; height: 28px; border-radius: 6px; }
	#product-table td.ms-act a:hover { background: #e8f3f0; }
	#product-table tr.ms-low td.ms-end { color: #a8322a; }
	.ms-flag { display: inline-block; margin-left: 6px; padding: 1px 7px; border-radius: 999px; font-size: 10.5px; font-weight: 600; background: #fde8e6; color: #a8322a; vertical-align: 1px; }
	.ms-empty { padding: 28px !important; text-align: center; color: #5d6d69; white-space: normal !important; }
	@media (max-width: 640px) {
		#wrapper { padding: 12px 16px; }
		.ms-title { font-size: 20px; }
		#ms-summary { grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 8px; }
		#ms-summary .fbt-card { padding: 10px 12px; }
		#ms-summary .fbt-card-value { font-size: 17px; overflow-wrap: anywhere; }
		.ms-panel { padding: 10px 10px 2px; }
	}
	@media print {
		body { background: #fff; }
		.ms-actions { display: none !important; }
		.ms-head, .ms-panel { border: 0; box-shadow: none; padding: 0; }
		.ms-scroll { overflow: visible; max-height: none; border: 0; }
		#product-table thead tr:last-child th, #product-table tfoot td { position: static; }
		#product-table td.ms-act, #product-table th.ms-act { display: none; }
		.ms-flag { background: none; padding: 0; }
	}
</style>
</head>
<%
dim filter1,filter2
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

'---- Parameter "kosong": halaman pemanggil kadang mengirim teks undefined / null
Function NoVal(v)
	Dim s
	s = LCase(Trim(v & ""))
	NoVal = (s = "" Or s = "undefined" Or s = "null" Or s = "-")
End Function

'---- Angka 2 desimal; nol ditampilkan abu-abu
Function NumCell(v, cls)
	Dim n
	n = cdbl(to_nol(v))
	If n = 0 Then
		NumCell = "<td class=""ms-num ms-zero"">" & FormatNumber(n, 2) & "</td>"
	Else
		NumCell = "<td class=""" & cls & """>" & FormatNumber(n, 2) & "</td>"
	End If
End Function

'---- Chip filter di kop laporan
Function Chip(lbl, v)
	Chip = ""
	If Not NoVal(v) Then Chip = "<span class=""ms-chip""><b>" & lbl & "</b>" & H(v) & "</span>"
End Function

dt 		 = request("dt")
pcompany = request("company")
pfcba 	 = request("fcba")
ptype 	 = request("type")
pmonth 	 = request("month")
pyear 	 = request("year")
psrc 	 = request("src")
pstore   = request("storecode")
psubtype = request("sub")

'---- Perbaikan: fcba=null / kosong dari halaman pemanggil -> pakai BA user yang login
if NoVal(pfcba) then pfcba = Session("fcba")
'---- Parameter yang disambung ke SQL di-escape
pfcba  = SqlStr(pfcba)
pmonth = SqlStr(pmonth)
pyear  = SqlStr(pyear)

q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & pfcba & "' "
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

'---- Perbaikan: dulu filter group dipakai bila type <> "-", sehingga type kosong
'---- mencari group '' dan laporan selalu kosong. Sekarang hanya bila group dipilih.
if not NoVal(ptype) then
		filter1 = " AND stock_api.get_stockgroup(st.fccode,st.fcba) ='" & SqlStr(ptype) & "' "
end if

if not NoVal(pstore) then
		filter2 = " AND s.storecode ='" & SqlStr(pstore) & "' "
end if


txtSQL = " SELECT " & _
       " st.fcba, " & _
       " s.storecode, " & _
       " st.fccode, " & _
       " stock_api.get_stockgroup(st.fccode,st.fcba) StockGroupcode, " & _
       " stockgroup_api.get_stockgrpname(stock_api.get_stockgroup(st.fccode,st.fcba),st.fcba) stockgroupname, " & _
       " nvl(stocksubgroup_api.get_stocksubgrpname(stock_api.get_StockSubGroup(st.fccode,st.fcba)),'-') stocksubgroupname, " & _
       " stock_api.get_ActiveMatrial(st.fccode,st.fcba) ActMat , " & _
       " stock_api.get_stockname(st.fccode, st.fcba) fcname,  " & _
       " stock_api.get_uom(st.fccode, st.fcba) unitofmeasurement,  " & _
       " (Nvl(s.openingbalanceqty,0)) openingbalanceqty, " & _
       " (Nvl(s.qtyreceived,0)) qtyreceived, " & _
       " (Nvl(s.qtyreceivedtask,0)) qtyreceivedtask, " & _
       " (Nvl(s.qtyreceivetaskdiv,0)) qtyreceivetaskdiv, " & _
       " (Nvl(s.qtytransfered,0)) qtytransfered, " & _
       " (Nvl(s.qtytransporttask,0)) qtytransporttask, " & _
       " (Nvl(s.qtytransporttaskdiv,0)) qtytransporttaskdiv, " & _
       " (Nvl(s.qtyreturned,0)) qtyreturned, " & _
       " (Nvl(s.qtyissued,0)) qtyissued, " & _
       " (Nvl(s.qtywriteoff,0)) qtywriteoff, " & _
       " (Nvl(s.endingbalanceqty,0)) endingbalance, " & _
       " Round(Nvl(s.amtperunit,0),2) amtperunit, " & _
       " Decode(s.storecode,'ST1',stock_api.get_stockminimum(st.fccode, st.fcba),0) get_stockminimum, " & _
       " SUM(Nvl(pd.requestedqty,0)) requestedqty, " & _
       " SUM(Nvl(pd.approvedqty,0)) approvedqty, " & _
       " SUM(Nvl(pd.orderedqty,0)) orderedty " & _
       " FROM " & _
       " stock st " & _
       " Left Join stockledger s on st.FCCODE=s.ITEMCODE and st.fcba=s.fcba" & _
       " Left Join PURCHASEREQUESTDETAILS pd on s.ITEMCODE = pd.ITEMCODE and s.FCBA = pd.FCBA " & _
       " WHERE 1=1  " & _
       " AND s.YEAR='" & pyear & "' AND s.MONTH='" & pmonth & "' AND st.FCBA ='" & pfcba & "' " & _
	   filter1 & _
	   filter2 & _
	   " GROUP BY  st.fcba, s.storecode, st.fccode, stock_api.get_stockgroup(st.fccode,st.fcba) , " & _
	   " stockgroup_api.get_stockgrpname(stock_api.get_stockgroup(st.fccode,st.fcba),st.fcba) , " & _
       " nvl(stocksubgroup_api.get_stocksubgrpname(stock_api.get_StockSubGroup(st.fccode,st.fcba)),'-') , " & _
       " stock_api.get_ActiveMatrial(st.fccode,st.fcba)  , stock_api.get_stockname(st.fccode, st.fcba) , " & _
       " stock_api.get_uom(st.fccode, st.fcba),Round(Nvl(s.amtperunit,0),2), " & _
	   " (Nvl(s.openingbalanceqty,0)), " & _
       " (Nvl(s.qtyreceived,0)), " & _
       " (Nvl(s.qtyreceivedtask,0)), " & _
       " (Nvl(s.qtyreceivetaskdiv,0)), " & _
       " (Nvl(s.qtytransfered,0)), " & _
       " (Nvl(s.qtytransporttask,0)), " & _
       " (Nvl(s.qtytransporttaskdiv,0)), " & _
       " (Nvl(s.qtyreturned,0)), " & _
       " (Nvl(s.qtyissued,0)), " & _
       " (Nvl(s.qtywriteoff,0)), " & _
       " (Nvl(s.endingbalanceqty,0)) " & _
       " Order by st.fcba,st.fccode,s.storecode"

txtSelect = pmonth & "/" & pyear

if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
end if
' response.Write txtsql
'response.Flush()
%>
	<body>
		<div id="wrapper">
		<%
		Set rsCat = connStr.Execute(txtSQL)
		%>
		<div class="ms-head">
			<div>
				<div class="ms-company">PT. <%=H(CompanyName)%></div>
				<div class="ms-ba"><%=H(FcbaName)%></div>
				<h2 class="ms-title">Monitoring Stock</h2>
				<div class="ms-chips">
					<span class="ms-chip"><b>Periode</b><%=H(request("month"))%>/<%=H(request("year"))%></span>
					<%=Chip("Store", request("storecode"))%>
					<%=Chip("Group", request("type"))%>
					<%=Chip("Sub Group", request("sub"))%>
				</div>
			</div>
			<div class="ms-actions">
				<button type="button" class="ms-print" onclick="window.print()">Cetak</button>
				<!--#include file="print_excel.asp"-->
			</div>
		</div>
		<div class="fbt-summary" id="ms-summary"></div>
		<div class="ms-panel">
		<div id="ms-tools"></div>
		<div class="ms-scroll">
		<table border="0" width="100%" id="product-table">
			<thead>
			<tr class="ms-xls">
				<th colspan="21" style=" background-color:#FFFFFF; color:#000000; border:none">
					<h4>PT. <%=H(CompanyName)%></h4>
					<h5><%=H(FcbaName)%></h5>
					<h2>MONITORING STOCK</h2>
				</th>
			</tr>
			<tr class="ms-xls">
				<th colspan="21" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=H(txtSelect)%></th>
			</tr>
			<tr>
				<th class="ms-act" data-nosort></th>
				<th>Store</th>
				<th>Stock Code</th>
				<th>Stock Name</th>
				<th>Stock Group</th>
				<th>Satuan</th>
				<th>Minimal Stock</th>
				<th>Qty Open</th>
				<th>Qty Received</th>
				<th>Qty RCV. Task</th>
				<th>Qty.Rcv.Task.Div</th>
				<th>Qty Issue</th>
				<th>Qty. Transport.Task</th>
				<th>Qty Transport Task Div</th>
				<th>Qty Adjustment</th>
				<th>Qty Return</th>
				<th>Qty End</th>
				<th>Qty PP</th>
				<th>Qty PP Approved</th>
				<th>Qty PP PO</th>
				<th>Price</th>
			</tr>
			</thead>
			<tbody>
				<%
				Dim v_qtyopen,v_qtyreceived,v_qtyreceivedtask,v_qtyreceivedtaskdiv,v_qtyissued,v_qtytransporttask,v_qtytransporttaskdiv,v_qtywriteoff,v_ending
				dim v_qtypr,v_qtyapprovedpr,v_qtypo,v_return,qtyEnd,qtyMin,rowCls

				v_qtyopen = 0
				v_qtyreceived = 0
				v_qtyreceivedtask = 0
				v_qtyreceivedtaskdiv = 0
				v_qtyissued = 0
				v_qtytransporttask = 0
				v_qtytransporttaskdiv = 0
				v_qtywriteoff = 0
				v_return = 0
				v_ending = 0
				v_qtypr = 0
				v_qtyapprovedpr = 0
				v_qtypo = 0
				hitungTambah = 0
				hitungKurang = 0
				no = 1
				While Not rsCat.EOF
					hitungTambah = cdbl(rsCat("openingbalanceqty")) + cdbl(rsCat("qtyreceived")) + cdbl(rsCat("qtyreceivedtask")) + cdbl(rsCat("qtyreceivetaskdiv"))+ cdbl(rsCat("qtywriteoff"))
					hitungKurang = cdbl(rsCat("qtyissued")) + cdbl(rsCat("qtytransporttask")) + cdbl(rsCat("qtytransporttaskdiv"))  + cdbl(rsCat("qtyreturned"))
					qtyEnd = hitungTambah - hitungKurang
					qtyMin = cdbl(to_nol(rsCat("get_stockminimum")))
					rowCls = ""
					if qtyMin > 0 and qtyEnd < qtyMin then rowCls = " class=""ms-low"""
				%>
				<tr<%=rowCls%>>
					<td class="ms-act"><a target="_blank" title="Lihat kartu stock" href="mam_data_view.Asp?stockcode=<%=U(rsCat("fccode"))%>&storecode=<%=U(rsCat("storecode"))%>&fcba=<%=U(rsCat("fcba"))%>&month=<%=U(pmonth)%>&year=<%=U(pyear)%>"><img src="IMAGES/view.png" alt="View Data" border="0"></a></td>
					<td class="ms-muted"><%=H(rsCat("storecode"))%></td>
					<td class="ms-code"><%=H(rsCat("fccode"))%></td>
					<td class="ms-wrap"><%=H(rsCat("fcname"))%></td>
					<td><%=H(rsCat("stocksubgroupname"))%></td>
					<td><%=H(rsCat("unitofmeasurement"))%></td>
					<%=NumCell(rsCat("get_stockminimum"), "ms-num")%>
					<%=NumCell(rsCat("openingbalanceqty"), "ms-num")%>
					<%=NumCell(rsCat("qtyreceived"), "ms-num")%>
					<%=NumCell(rsCat("qtyreceivedtask"), "ms-num")%>
					<%=NumCell(rsCat("qtyreceivetaskdiv"), "ms-num")%>
					<%=NumCell(rsCat("qtyissued"), "ms-num")%>
					<%=NumCell(rsCat("qtytransporttask"), "ms-num")%>
					<%=NumCell(rsCat("qtytransporttaskdiv"), "ms-num")%>
					<%=NumCell(rsCat("qtywriteoff"), "ms-num")%>
					<%=NumCell(rsCat("qtyreturned"), "ms-num")%>
					<td class="ms-end"><%=FormatNumber(qtyEnd,2)%><%if rowCls <> "" then%><span class="ms-flag" title="Di bawah minimal stock">min</span><%end if%></td>
					<%=NumCell(rsCat("requestedqty"), "ms-num")%>
					<%=NumCell(rsCat("approvedqty"), "ms-num")%>
					<%=NumCell(rsCat("orderedty"), "ms-num")%>
					<%=NumCell(rsCat("amtperunit"), "ms-num")%>
				</tr>
				<%
					v_qtyopen = v_qtyopen + cdbl(rsCat("openingbalanceqty"))
					v_qtyreceived = v_qtyreceived + cdbl(rsCat("qtyreceived"))
					v_qtyreceivedtask = v_qtyreceivedtask + cdbl(rsCat("qtyreceivedtask"))
					v_qtyreceivedtaskdiv = v_qtyreceivedtaskdiv + cdbl(rsCat("qtyreceivetaskdiv"))
					v_qtyissued = v_qtyissued + cdbl(rsCat("qtyissued"))
					v_qtytransporttask = v_qtytransporttask + cdbl(rsCat("qtytransporttask"))
					v_qtytransporttaskdiv = v_qtytransporttaskdiv  + cdbl(rsCat("qtytransporttaskdiv"))
					v_qtywriteoff = v_qtywriteoff + cdbl(rsCat("qtywriteoff"))
					v_return = v_return + cdbl(rsCat("qtyreturned"))
					'---- Perbaikan: total Qty End = jumlah Qty End yang tampil di baris (dulu memakai endingbalance DB)
					v_ending = v_ending + qtyEnd
					v_qtypr = v_qtypr + cdbl(rsCat("requestedqty"))
					v_qtyapprovedpr = v_qtyapprovedpr + cdbl(rsCat("approvedqty"))
					v_qtypo = v_qtypo + cdbl(rsCat("orderedty"))
					rsCat.MoveNext
					'response.Flush()
					no=no+1
				Wend
				rsCat.close()
				Set rsCat = Nothing
				if no = 1 then%>
				<tr><td colspan="21" class="ms-empty">Tidak ada data stock untuk periode / filter ini.</td></tr>
				<%end if%>
			</tbody>
			<tfoot>
			<tr bgcolor="#F4F4A0">
				<td colspan="7">Total (<span class="fbt-count"><%=no-1%></span> item)</td>
				<td class="ms-num" data-sum="7"><%=FormatNumber(v_qtyopen,2)%></td>
				<td class="ms-num" data-sum="8"><%=FormatNumber(v_qtyreceived,2)%></td>
				<td class="ms-num" data-sum="9"><%=FormatNumber(v_qtyreceivedtask,2)%></td>
				<td class="ms-num" data-sum="10"><%=FormatNumber(v_qtyreceivedtaskdiv,2)%></td>
				<td class="ms-num" data-sum="11"><%=FormatNumber(v_qtyissued,2)%></td>
				<td class="ms-num" data-sum="12"><%=FormatNumber(v_qtytransporttask,2)%></td>
				<td class="ms-num" data-sum="13"><%=FormatNumber(v_qtytransporttaskdiv,2)%></td>
				<td class="ms-num" data-sum="14"><%=FormatNumber(v_qtywriteoff,2)%></td>
				<td class="ms-num" data-sum="15"><%=FormatNumber(v_return,2)%></td>
				<td class="ms-num" data-sum="16"><%=FormatNumber(v_ending,2)%></td>
				<td class="ms-num" data-sum="17"><%=FormatNumber(v_qtypr,2)%></td>
				<td class="ms-num" data-sum="18"><%=FormatNumber(v_qtyapprovedpr,2)%></td>
				<td class="ms-num" data-sum="19"><%=FormatNumber(v_qtypo,2)%></td>
				<td></td>
			</tr>
			</tfoot>
		</table>
		</div>
		<div id="ms-foot"></div>
		</div>
		<input  type="hidden" value="<%=H(txtTable)%>" name="txtTable">
		<input  type="hidden" value="<%=H(txtModule)%>" name="txtModule">
		</div>
	<script type="text/javascript">
	// Tabel ala DataTables + kartu ringkasan (hanya tampilan; angka dari baris yang dirender server)
	window.addEventListener('load', function () {
		var table = document.getElementById('product-table');
		var summary = document.getElementById('ms-summary');
		if (!window.FbpTable || !table) return;   // tetap dipasang walau data kosong: toolbar cari & jumlah baris tetap tampil
		function fmt(n, d) { return n.toFixed(d || 0).replace(/\B(?=(\d{3})+(?!\d))/g, ','); }
		function txt(r, i) { return (r.cells[i].textContent || '').replace(/\s+/g, ' ').trim(); }
		function num(r, i) { return parseFloat(txt(r, i).replace(/,/g, '').replace(/[^\d.\-]/g, '')) || 0; }
		function card(label, value, note, extra) {
			return '<div class="fbt-card' + (extra || '') + '"><div class="fbt-card-label">' + label + '</div>' +
				'<div class="fbt-card-value">' + value + '</div><div class="fbt-card-note">' + note + '</div></div>';
		}
		FbpTable.patchExport();
		FbpTable.init(table, {
			pageSize: 20,
			onChange: function (rows) {
				var s = {}, store = {}, low = 0, idle = 0, prOpen = 0, value = 0;
				for (var c = 7; c <= 19; c++) s[c] = 0;
				rows.forEach(function (r) {
					for (var c = 7; c <= 19; c++) s[c] += num(r, c);
					store[txt(r, 1) || '-'] = 1;
					if (r.classList.contains('ms-low')) low++;
					var moved = 0;
					for (var m = 8; m <= 15; m++) moved += Math.abs(num(r, m));
					if (!moved) idle++;
					if (num(r, 18) > num(r, 19)) prOpen++;
					value += num(r, 16) * num(r, 20);
				});
				Array.prototype.forEach.call(table.tFoot.querySelectorAll('[data-sum]'), function (td) {
					td.textContent = fmt(s[td.getAttribute('data-sum')], 2);
				});
				summary.innerHTML =
					card('Item stock', fmt(rows.length), Object.keys(store).length + ' store') +
					card('Nilai stock akhir', fmt(value), 'qty end &times; price') +
					card('Di bawah minimal', fmt(low), 'qty end &lt; minimal stock', low ? ' warn' : '') +
					card('Tanpa pergerakan', fmt(idle), 'tidak ada masuk/keluar bulan ini') +
					card('PP belum di-PO', fmt(prOpen), 'qty PP approved &gt; qty PP PO', prOpen ? ' accent' : '');
			}
		});
		// Toolbar & pager di luar area gulir agar tetap terlihat walau tabel lebar
		var tools = table.querySelector('.fbt-toolbar-row .fbt-tools'), foot = table.parentNode.querySelector('.fbt-foot');
		if (tools) document.getElementById('ms-tools').appendChild(tools);
		if (foot) document.getElementById('ms-foot').appendChild(foot);
	});
	</script>
		<!--#include file="Bottom_Report.asp"-->
