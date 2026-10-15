<!--#include file="txtFunction.inc"-->

<style>
/* Di-scope ke .mdi agar tidak memengaruhi tabel lain di halaman */
.mdi {
	--brand: #1B7863;
	--ink: #17262a;
	--muted: #5d6d69;
	--line: #e1e7e5;
	margin-top: 8px;
	color: var(--ink);
	font-family: 'Plus Jakarta Sans', 'Segoe UI', Roboto, Arial, sans-serif;
	font-size: 14px;
}
.mdi-section-title {
	display: flex; align-items: center; gap: 10px;
	margin: 8px 0 12px;
	font-size: 12px; font-weight: 700; letter-spacing: 1px; text-transform: uppercase;
	color: var(--muted);
}
.mdi-section-title::after { content: ""; flex: 1; height: 1px; background: var(--line); }
.mdi-card {
	background: #fff;
	border: 1px solid var(--line);
	border-radius: 14px;
	padding: 18px 20px 20px;
}
.mdi-head {
	display: flex; flex-wrap: wrap; align-items: flex-start; justify-content: space-between; gap: 12px;
	margin-bottom: 14px;
}
.mdi-head h4 { margin: 0; font-family: inherit; font-size: 15px; font-weight: 700; color: var(--ink); }
.mdi-head p { margin: 2px 0 0; font-size: 12px; color: var(--muted); }
.mdi-date {
	display: inline-flex; align-items: center; gap: 8px;
	padding: 6px 12px;
	font-size: 13px; font-weight: 600;
	background: #f3f6f5; border: 1px solid var(--line); border-radius: 999px;
}
.mdi-date .fa { color: var(--brand); }
.mdi-legend { display: flex; flex-wrap: wrap; gap: 14px; margin-top: 12px; font-size: 12px; color: var(--muted); }
.mdi-legend span { display: inline-flex; align-items: center; gap: 6px; }
.mdi-legend i { width: 12px; height: 12px; border-radius: 3px; display: inline-block; }

.mdi-wrap { overflow: auto; border: 1px solid var(--line); border-radius: 10px; }
.mdi .table1 {
	width: 100%;
	border-collapse: collapse;
	font-size: 13px;
	color: var(--ink);
}
.mdi .table1 th {
	padding: 10px 12px;
	background: #f6f9f8;
	color: var(--muted);
	font-size: 11px; font-weight: 700; letter-spacing: .5px; text-transform: uppercase;
	text-align: center;
	border: 0; border-bottom: 1px solid var(--line);
	white-space: nowrap;
}
.mdi .table1 th.mdi-pt { text-align: left; }
.mdi .table1 td {
	padding: 6px 8px;
	border: 0; border-bottom: 1px solid #eef2f1;
	text-align: center;
	font-variant-numeric: tabular-nums;
}
.mdi .table1 tbody tr:last-child td { border-bottom: 0; }
.mdi .table1 tbody tr:hover td { background: #f8fbfa; }
.mdi .table1 td:first-child { text-align: left; padding-left: 12px; }
.mdi .table1 td a { color: var(--brand); font-weight: 600; }
.mdi .pill {
	display: inline-block;
	min-width: 72px;
	padding: 4px 10px;
	border-radius: 999px;
	font-weight: 600;
}
.mdi .pill-none { background: #fdecec; color: #b42318; }
.mdi .pill-part { background: #fff4d6; color: #8a5a00; }
.mdi .pill-full { background: #e6f4ea; color: #146c2e; }
.mdi .mdi-empty { padding: 24px; color: var(--muted); text-align: center; }
</style>

	<script src="/js/jquery-1.9.1.js"></script>

<script>
	$(function() {
		getMonitoring();
	});


    var timerMonitoringVar = setInterval(function() { getMonitoring() }, 3600000);

	// Sel persentase: 0 = belum input, 1-99 = sebagian, 100 = lengkap
	function pctCell(val) {
		var cls = 'pill-full';
		if (val == 0) cls = 'pill-none';
		else if (val > 0 && val < 100) cls = 'pill-part';
		return '<td><span class="pill ' + cls + '">' + numformat2(val) + '%</span></td>';
	}

	function getMonitoring()
	{			$('#bodytable').empty();
				$.post("GetMonitoringDataInputView.asp?validasi=" + 1, function(data) {
					var json  = JSON.parse(data);
					var jumlviewtbs  = json.dataview.length;

					for (var i = 0; i < jumlviewtbs; i++) {
						var day_4   = json.dataview[i].hmin4_point;
						var day_3 	= json.dataview[i].hmin3_point;
						var day_2 	= json.dataview[i].hmin2_point;
						var day_1 	= json.dataview[i].hmin1_point;
						if (jumlviewtbs>0){
							var docketlist = '';

							docketlist += '<tr>';
							docketlist += '<td><a target=_blank href="MonitoringDataInputDet.Asp?company='+json.dataview[i].companycode+ '">'+json.dataview[i].companycode+'</a></td>';
							docketlist += pctCell(day_4);
							docketlist += pctCell(day_3);
							docketlist += pctCell(day_2);
							docketlist += pctCell(day_1);
							docketlist += '</tr>';
							$('#bodytable').append(docketlist);
						}

					}
					if (jumlviewtbs == 0) {
						$('#bodytable').append('<tr><td colspan="5" class="mdi-empty">Data Tidak Ada</td></tr>');
					}
				});

	}



	// Format tampilan tabel: ribuan koma + 2 desimal (1,234,567.89). numformat() lama tetap dipakai untuk input.
	function numformat2(val) {
		var n = parseFloat(String(val == null ? '' : val).replace(/,/g, ''));
		if (isNaN(n)) return val == null ? '' : val;
		var p = (Math.abs(n) < 0.005 ? 0 : n).toFixed(2).split('.');
		return p[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',') + '.' + p[1];
	}

	function numformat(val) {
		var nStr = val.toString();
		nStr = nStr.replace(/\,/g, "");
		var x = nStr.split('.');
		var x1 = x[0];
		var x2 = x.length > 1 ? '.' + x[1] : '';
		var rgx = /(\d+)(\d{3})/;
		while (rgx.test(x1)) {
			x1 = x1.replace(rgx, '$1' + '.' + '$2');
		}
		return x1 + x2;
	}


</script>
<%
dim tglHariIni,fcbasession
fcbasession = Session("fcBA")
tglHariIni = textdate(Now)

%>
<div class="mdi">
	<div class="mdi-section-title">Input Data</div>
	<div class="mdi-card">
		<div class="mdi-head">
			<div>
				<h4>Monitoring Data Inputan</h4>
				<p>Persentase kelengkapan input data per PT, 4 hari terakhir. Klik kode PT untuk melihat detail.</p>
			</div>
			<span class="mdi-date"><i class="fa fa-calendar" aria-hidden="true"></i> Hari ini: <%=tglHariIni%></span>
		</div>
		<div class="mdi-wrap">
			<table class="table1">
				<thead>
				 <tr>
					<th class="mdi-pt" rowspan="2">PT</th>
					<th colspan="4">Tanggal</th>
				 </tr>
				 <tr>
					<th><%=textdate(Now-4)%></th>
					<th><%=textdate(Now-3)%></th>
					<th><%=textdate(Now-2)%></th>
					<th><%=textdate(Now-1)%></th>
				 </tr>
				</thead>
				<tbody id="bodytable">
				</tbody>
			</table>
		</div>
		<div class="mdi-legend">
			<span><i style="background:#fdecec; border:1px solid #f3c7c3"></i>0% &middot; belum input</span>
			<span><i style="background:#fff4d6; border:1px solid #f1dc9c"></i>1&ndash;99% &middot; sebagian</span>
			<span><i style="background:#e6f4ea; border:1px solid #b9dcc2"></i>100% &middot; lengkap</span>
		</div>
	</div>
</div>
