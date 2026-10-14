<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="include/db_param.asp"-->
<%
' =====================================================================
' Popup pilih barang untuk Kartu Stock.
' Dulu: semua barang (~9.000 baris, 5 MB) dicetak sekaligus -> 20+ detik.
' Sekarang: halaman ringan + data per halaman lewat AJAX (?ajax=1), dicari & dipaging di database.
' Kontrak ke halaman pembuka tidak berubah: isi #itemcode (value), #StockName & #productcode (teks).
' =====================================================================
Dim ksBa
ksBa = UCase(Trim(Request("fcba") & ""))

Function KsJson(v)
	Dim s, i, c, code, o
	s = "" & v
	o = ""
	For i = 1 To Len(s)
		c = Mid(s, i, 1)
		code = AscW(c)
		If code < 0 Then code = code + 65536
		Select Case code
			Case 34: o = o & "\"""
			Case 92: o = o & "\\"
			Case 10: o = o & "\n"
			Case 13: o = o & "\r"
			Case 9: o = o & "\t"
			Case Else
				If code < 32 Or code > 126 Then
					o = o & "\u" & Right("000" & Hex(code), 4)
				Else
					o = o & c
				End If
		End Select
	Next
	KsJson = """" & o & """"
End Function

If Request("ajax") = "1" Then
	Response.ContentType = "application/json"
	Response.CacheControl = "no-cache"
	Response.Expires = -1

	Dim ksStart, ksLen, ksQ, ksTokens, ksT, ksWhere, ksParams, ksSql, ksRs, ksOut, ksFirst, ksFiltered, ksTotal, ksK
	ksStart = 0 : ksLen = 25
	If IsNumeric(Request("start")) Then ksStart = CLng(Request("start"))
	If IsNumeric(Request("length")) Then ksLen = CLng(Request("length"))
	If ksStart < 0 Then ksStart = 0
	If ksLen <> 10 And ksLen <> 25 And ksLen <> 50 And ksLen <> 100 Then ksLen = 25

	' sumber data sama dengan query lama: barang kode 1xx di BA, grup stok aktif
	ksWhere = " FROM stock a JOIN stockgroup b ON SUBSTR(a.fccode, 1, 3) = b.fccode AND b.activation = 'Y' " & _
	          " WHERE a.fcba = ? AND SUBSTR(a.fccode, 1, 1) = '1' "
	ksParams = Array(ksBa)

	' cari: tiap kata harus ada di kode atau nama (maks. 5 kata); % dan _ diperlakukan sebagai huruf biasa
	ksQ = UCase(Trim(Request("q") & ""))
	If ksQ <> "" Then
		ksTokens = Split(ksQ, " ")
		ksK = 0
		For Each ksT In ksTokens
			If ksT <> "" And ksK < 5 Then
				ksT = "%" & Replace(Replace(Replace(ksT, "\", "\\"), "%", "\%"), "_", "\_") & "%"
				ksWhere = ksWhere & " AND (UPPER(a.fccode) LIKE ? ESCAPE '\' OR UPPER(a.fcname) LIKE ? ESCAPE '\') "
				ReDim Preserve ksParams(UBound(ksParams) + 2)
				ksParams(UBound(ksParams) - 1) = ksT
				ksParams(UBound(ksParams)) = ksT
				ksK = ksK + 1
			End If
		Next
	End If

	' satu query: halaman yang diminta + jumlah baris sesuai filter (COUNT(*) OVER)
	ksSql = "SELECT itemcode, itemname, cnt FROM (" & _
	        " SELECT x.*, ROWNUM rn FROM (" & _
	        "  SELECT a.fccode itemcode, a.fcname itemname, COUNT(*) OVER () cnt" & ksWhere & _
	        "  GROUP BY a.fccode, a.fcname ORDER BY a.fccode, a.fcname" & _
	        " ) x WHERE ROWNUM <= " & (ksStart + ksLen) & _
	        ") WHERE rn > " & ksStart
	Set ksRs = DbQuery(ksSql, ksParams)

	ksFiltered = -1
	ksOut = ""
	ksFirst = True
	Do While Not ksRs.EOF
		If ksFiltered < 0 Then ksFiltered = CLng(ksRs("cnt"))
		If Not ksFirst Then ksOut = ksOut & ","
		ksFirst = False
		ksOut = ksOut & "[" & KsJson(ksRs("itemcode")) & "," & KsJson(ksRs("itemname")) & "]"
		ksRs.MoveNext
	Loop
	ksRs.Close
	Set ksRs = Nothing

	' halaman di luar jangkauan (mis. setelah filter) -> jumlah dihitung terpisah
	If ksFiltered < 0 Then
		Set ksRs = DbQuery("SELECT COUNT(*) n FROM (SELECT a.fccode, a.fcname" & ksWhere & " GROUP BY a.fccode, a.fcname)", ksParams)
		ksFiltered = CLng(ksRs("n"))
		ksRs.Close
	End If

	' total tanpa filter: disimpan di sesi per BA agar pencarian tidak menghitung ulang semuanya
	If ksQ = "" Then
		ksTotal = ksFiltered
		Session("ks_total_" & ksBa) = ksTotal
	ElseIf IsNumeric(Session("ks_total_" & ksBa)) And Session("ks_total_" & ksBa) & "" <> "" Then
		ksTotal = CLng(Session("ks_total_" & ksBa))
	Else
		Set ksRs = DbQuery("SELECT COUNT(*) n FROM (SELECT a.fccode, a.fcname FROM stock a JOIN stockgroup b ON SUBSTR(a.fccode, 1, 3) = b.fccode AND b.activation = 'Y' WHERE a.fcba = ? AND SUBSTR(a.fccode, 1, 1) = '1' GROUP BY a.fccode, a.fcname)", Array(ksBa))
		ksTotal = CLng(ksRs("n"))
		ksRs.Close
		Session("ks_total_" & ksBa) = ksTotal
	End If

	Response.Write "{""total"":" & ksTotal & ",""filtered"":" & ksFiltered & ",""start"":" & ksStart & ",""length"":" & ksLen & ",""rows"":[" & ksOut & "]}"
	Response.End
End If
%>
<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>Pilih Barang &middot; Kartu Stock</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
	:root {
		--brand: #1B7863; --brand-dark: #12584a; --tint: #eef6f3; --tint-2: #e2f4ec;
		--ink: #17262a; --muted: #5d6d69; --line: #e1e7e5; --bg: #f3f6f5;
	}
	* { box-sizing: border-box; }
	html, body { height: 100%; }
	body { margin: 0; font-family: 'Plus Jakarta Sans', 'Segoe UI', Roboto, Arial, sans-serif; font-size: 13px; color: var(--ink); background: var(--bg); }
	.ks { display: flex; flex-direction: column; height: 100%; max-width: 1100px; margin: 0 auto; padding: 16px; gap: 12px; }

	.ks-head { display: flex; align-items: center; gap: 14px; padding: 14px 18px; color: #fff; border-radius: 14px;
		background: linear-gradient(120deg, var(--brand) 0%, var(--brand-dark) 100%); }
	.ks-head h1 { margin: 0; font-size: 17px; font-weight: 700; }
	.ks-head p { margin: 2px 0 0; font-size: 12px; opacity: .85; }
	.ks-head .ks-ba { margin-left: auto; padding: 5px 12px; font-size: 12px; font-weight: 600; border-radius: 999px; background: rgba(255,255,255,.14); border: 1px solid rgba(255,255,255,.22); white-space: nowrap; }
	.ks-x { width: 34px; height: 34px; display: grid; place-items: center; color: #fff; background: rgba(255,255,255,.12); border: 0; border-radius: 10px; cursor: pointer; font-size: 18px; line-height: 1; }
	.ks-x:hover { background: rgba(255,255,255,.24); }

	.ks-card { display: flex; flex-direction: column; flex: 1 1 auto; min-height: 0; background: #fff; border: 1px solid var(--line); border-radius: 14px; overflow: hidden; }
	.ks-tools { display: flex; flex-wrap: wrap; align-items: center; gap: 10px 14px; padding: 14px 16px; border-bottom: 1px solid var(--line); }
	.ks-search { position: relative; flex: 1 1 320px; }
	.ks-search input { width: 100%; height: 40px; padding: 0 38px 0 38px; font: inherit; font-size: 14px; color: var(--ink); background: #f6f9f8; border: 1px solid var(--line); border-radius: 10px; outline: none; }
	.ks-search input:focus { background: #fff; border-color: var(--brand); box-shadow: 0 0 0 3px rgba(27,120,99,.18); }
	.ks-search svg { position: absolute; left: 12px; top: 50%; transform: translateY(-50%); color: #8a9a96; pointer-events: none; }
	.ks-clear { position: absolute; right: 6px; top: 50%; transform: translateY(-50%); width: 28px; height: 28px; border: 0; border-radius: 8px; background: transparent; color: var(--muted); cursor: pointer; font-size: 16px; display: none; }
	.ks-clear:hover { background: var(--tint); }
	.ks-size { display: flex; align-items: center; gap: 8px; color: var(--muted); white-space: nowrap; }
	.ks-size select { height: 40px; padding: 0 10px; font: inherit; color: var(--ink); border: 1px solid var(--line); border-radius: 10px; background: #fff; }
	.ks-hint { width: 100%; margin: -2px 0 0; font-size: 11.5px; color: var(--muted); }
	.ks-hint kbd { padding: 0 5px; font: 600 11px inherit; border: 1px solid var(--line); border-bottom-width: 2px; border-radius: 4px; background: #fff; }

	.ks-scroll { position: relative; flex: 1 1 auto; min-height: 200px; overflow: auto; }
	table { width: 100%; border-collapse: separate; border-spacing: 0; }
	thead th { position: sticky; top: 0; z-index: 1; padding: 10px 16px; font-size: 11px; font-weight: 700; letter-spacing: .4px; text-transform: uppercase; text-align: left; color: var(--muted); background: #f6f9f8; border-bottom: 1px solid var(--line); }
	thead th.act { width: 110px; text-align: right; }
	tbody td { padding: 9px 16px; border-bottom: 1px solid #edf1f0; vertical-align: middle; }
	tbody tr { cursor: pointer; }
	tbody tr:hover td { background: #f6faf8; }
	tbody tr.active td { background: var(--tint-2); }
	tbody tr.active td:first-child { box-shadow: inset 3px 0 0 var(--brand); }
	.code { display: inline-block; padding: 2px 8px; font: 600 12.5px ui-monospace, Consolas, monospace; color: var(--brand-dark); background: var(--tint); border-radius: 6px; white-space: nowrap; }
	.name { font-size: 13.5px; }
	mark { padding: 0 1px; color: inherit; background: #ffe58a; border-radius: 3px; }
	td.act { text-align: right; white-space: nowrap; }
	.ks-pick { height: 32px; padding: 0 14px; font: 600 12.5px inherit; color: #fff; background: var(--brand); border: 0; border-radius: 8px; cursor: pointer; }
	.ks-pick:hover { background: var(--brand-dark); }
	.ks-pick:focus-visible { outline: 2px solid var(--brand); outline-offset: 2px; }

	.ks-state { padding: 46px 16px; text-align: center; color: var(--muted); }
	.ks-state b { display: block; margin-bottom: 4px; font-size: 14px; color: var(--ink); }
	.sk td span { display: block; height: 12px; border-radius: 6px; background: linear-gradient(90deg, #eef2f1 25%, #f6f9f8 37%, #eef2f1 63%); background-size: 400% 100%; animation: sk 1.2s ease infinite; }
	@keyframes sk { 0% { background-position: 100% 50%; } 100% { background-position: 0 50%; } }
	.ks-busy::after { content: ""; position: absolute; inset: 0; background: rgba(255,255,255,.55); }

	.ks-foot { display: flex; flex-wrap: wrap; align-items: center; justify-content: space-between; gap: 10px; padding: 12px 16px; border-top: 1px solid var(--line); }
	.ks-info { color: var(--muted); }
	.ks-info b { color: var(--ink); }
	.ks-pager { display: flex; flex-wrap: wrap; gap: 4px; }
	.ks-pager button { min-width: 34px; height: 34px; padding: 0 10px; font: 600 12.5px inherit; color: var(--ink); background: #fff; border: 1px solid var(--line); border-radius: 8px; cursor: pointer; }
	.ks-pager button:hover:not(:disabled) { background: var(--tint); border-color: #cfe0da; }
	.ks-pager button[aria-current="page"] { color: #fff; background: var(--brand); border-color: var(--brand); }
	.ks-pager button:disabled { opacity: .45; cursor: default; }
	.ks-pager span { align-self: center; padding: 0 4px; color: var(--muted); }

	.ks-toast { position: fixed; left: 50%; bottom: 20px; transform: translateX(-50%); padding: 10px 16px; color: #fff; background: #17262a; border-radius: 10px; font-size: 13px; display: none; }

	@media (max-width: 640px) {
		.ks { padding: 10px; }
		.ks-head .ks-ba { display: none; }
		thead th.act, td.act { width: auto; }
		.ks-pick { padding: 0 10px; }
		.ks-foot { justify-content: center; }
	}
</style>
</head>
<body>
<div class="ks">
	<header class="ks-head">
		<div>
			<h1>Pilih Barang</h1>
			<p>Kartu Stock &middot; klik baris atau tekan Enter untuk memilih</p>
		</div>
		<span class="ks-ba">BA <%=Server.HTMLEncode(ksBa)%></span>
		<button type="button" class="ks-x" id="ksClose" title="Tutup (Esc)" aria-label="Tutup">&times;</button>
	</header>

	<section class="ks-card">
		<div class="ks-tools">
			<div class="ks-search">
				<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" aria-hidden="true"><circle cx="11" cy="11" r="7"/><path d="m20 20-4-4"/></svg>
				<input type="search" id="ksQ" placeholder="Cari kode atau nama barang..." autocomplete="off" aria-label="Cari kode atau nama barang" autofocus>
				<button type="button" class="ks-clear" id="ksClear" title="Hapus pencarian" aria-label="Hapus pencarian">&times;</button>
			</div>
			<label class="ks-size">Tampilkan
				<select id="ksLen" aria-label="Jumlah baris per halaman">
					<option value="10">10</option>
					<option value="25" selected>25</option>
					<option value="50">50</option>
					<option value="100">100</option>
				</select>
			</label>
			<p class="ks-hint"><kbd>&uarr;</kbd> <kbd>&darr;</kbd> pindah baris &middot; <kbd>Enter</kbd> pilih &middot; <kbd>PgUp</kbd> <kbd>PgDn</kbd> ganti halaman &middot; <kbd>Esc</kbd> tutup</p>
		</div>
		<div class="ks-scroll" id="ksScroll">
			<table>
				<thead><tr><th style="width:170px">Kode Barang</th><th>Nama Barang</th><th class="act"></th></tr></thead>
				<tbody id="ksBody"></tbody>
			</table>
		</div>
		<div class="ks-foot">
			<div class="ks-info" id="ksInfo">Memuat...</div>
			<nav class="ks-pager" id="ksPager" aria-label="Halaman"></nav>
		</div>
	</section>
</div>
<div class="ks-toast" id="ksToast" role="status"></div>

<script>
(function () {
	'use strict';
	var BA = <%=KsJson(ksBa)%>;
	var state = { start: 0, length: 25, q: '', filtered: 0, total: 0, rows: [], active: -1, req: 0 };
	var $ = function (id) { return document.getElementById(id); };
	var body = $('ksBody'), info = $('ksInfo'), pager = $('ksPager'), scroll = $('ksScroll'), q = $('ksQ');

	function fmt(n) { return String(n).replace(/\B(?=(\d{3})+(?!\d))/g, '.'); }
	function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
	function mark(s) {
		var t = esc(s), words = state.q.trim().split(/\s+/).filter(Boolean);
		if (!words.length) return t;
		var re = new RegExp('(' + words.map(function (w) { return esc(w).replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }).join('|') + ')', 'gi');
		return t.replace(re, '<mark>$1</mark>');
	}
	function toast(m) { var t = $('ksToast'); t.textContent = m; t.style.display = 'block'; clearTimeout(toast.t); toast.t = setTimeout(function () { t.style.display = 'none'; }, 3500); }

	function skeleton() {
		var h = '';
		for (var i = 0; i < Math.min(state.length, 12); i++) h += '<tr class="sk"><td><span style="width:110px"></span></td><td><span style="width:' + (40 + (i * 37) % 50) + '%"></span></td><td></td></tr>';
		body.innerHTML = h;
	}

	function load() {
		var id = ++state.req;
		if (!state.rows.length) skeleton(); else scroll.classList.add('ks-busy');
		var url = 'popup_kartustock.asp?ajax=1&fcba=' + encodeURIComponent(BA) + '&start=' + state.start + '&length=' + state.length + '&q=' + encodeURIComponent(state.q.trim());
		var x = new XMLHttpRequest();
		x.open('GET', url, true);
		x.setRequestHeader('X-Requested-With', 'XMLHttpRequest');
		x.onload = function () {
			if (id !== state.req) return;   // jawaban lama (pengguna sudah mengetik lagi)
			scroll.classList.remove('ks-busy');
			var d = null;
			try { d = JSON.parse(x.responseText); } catch (e) { /* bukan JSON */ }
			if (x.status === 401) { location.href = '/logout.asp'; return; }
			if (!d || !d.rows) { body.innerHTML = '<tr><td colspan="3" class="ks-state"><b>Data gagal dimuat</b>Coba muat ulang popup.</td></tr>'; info.textContent = ''; pager.innerHTML = ''; return; }
			state.rows = d.rows; state.filtered = d.filtered; state.total = d.total;
			render();
		};
		x.onerror = function () { if (id === state.req) { scroll.classList.remove('ks-busy'); toast('Koneksi gagal, coba lagi.'); } };
		x.send();
	}

	function render() {
		if (!state.rows.length) {
			body.innerHTML = '<tr><td colspan="3" class="ks-state"><b>Barang tidak ditemukan</b>' + (state.q ? 'Tidak ada kode/nama yang cocok dengan &ldquo;' + esc(state.q.trim()) + '&rdquo;.' : 'Belum ada data barang untuk BA ini.') + '</td></tr>';
		} else {
			body.innerHTML = state.rows.map(function (r, i) {
				return '<tr data-i="' + i + '"><td><span class="code">' + mark(r[0]) + '</span></td><td class="name">' + mark(r[1]) + '</td>' +
					'<td class="act"><button type="button" class="ks-pick" data-i="' + i + '">Pilih</button></td></tr>';
			}).join('');
		}
		var from = state.filtered ? state.start + 1 : 0, to = Math.min(state.start + state.length, state.filtered);
		info.innerHTML = state.filtered
			? 'Menampilkan <b>' + fmt(from) + '&ndash;' + fmt(to) + '</b> dari <b>' + fmt(state.filtered) + '</b> barang' + (state.q ? ' (difilter dari ' + fmt(state.total) + ')' : '')
			: (state.q ? '0 barang cocok (dari ' + fmt(state.total) + ')' : 'Tidak ada data');
		renderPager();
		setActive(state.rows.length ? 0 : -1, false);
		scroll.scrollTop = 0;
	}

	function renderPager() {
		var pages = Math.max(1, Math.ceil(state.filtered / state.length)), cur = Math.floor(state.start / state.length) + 1;
		var b = function (label, page, opt) {
			opt = opt || {};
			return '<button type="button" data-p="' + page + '"' + (opt.dis ? ' disabled' : '') + (opt.cur ? ' aria-current="page"' : '') + (opt.aria ? ' aria-label="' + opt.aria + '"' : '') + '>' + label + '</button>';
		};
		var h = b('&lsaquo;', cur - 1, { dis: cur <= 1, aria: 'Halaman sebelumnya' });
		var list = [];
		if (pages <= 7) { for (var p = 1; p <= pages; p++) list.push(p); }
		else {
			list.push(1);
			var a = Math.max(2, cur - 1), z = Math.min(pages - 1, cur + 1);
			if (cur <= 3) { a = 2; z = 4; } else if (cur >= pages - 2) { a = pages - 3; z = pages - 1; }
			if (a > 2) list.push(0);
			for (var k = a; k <= z; k++) list.push(k);
			if (z < pages - 1) list.push(0);
			list.push(pages);
		}
		list.forEach(function (n) { h += n === 0 ? '<span>&hellip;</span>' : b(fmt(n), n, { cur: n === cur, aria: 'Halaman ' + n }); });
		h += b('&rsaquo;', cur + 1, { dis: cur >= pages, aria: 'Halaman berikutnya' });
		pager.innerHTML = h;
	}

	function goPage(p) {
		var pages = Math.max(1, Math.ceil(state.filtered / state.length));
		p = Math.min(Math.max(1, p), pages);
		var ns = (p - 1) * state.length;
		if (ns === state.start) return;
		state.start = ns;
		load();
	}

	function setActive(i, focusScroll) {
		var trs = body.querySelectorAll('tr[data-i]');
		if (state.active >= 0 && trs[state.active]) trs[state.active].classList.remove('active');
		state.active = i;
		if (i >= 0 && trs[i]) {
			trs[i].classList.add('active');
			if (focusScroll !== false) trs[i].scrollIntoView({ block: 'nearest' });
		}
	}

	function choose(i) {
		var r = state.rows[i];
		if (!r) return;
		var op = null;
		try { op = window.opener && !window.opener.closed ? window.opener.parent.document : null; } catch (e) { op = null; }
		if (!op || !op.getElementById('itemcode')) { toast('Halaman pemanggil tidak ditemukan. Kode terpilih: ' + r[0]); return; }
		op.getElementById('itemcode').value = r[0];
		var sn = op.getElementById('StockName'), pc = op.getElementById('productcode');
		if (sn) sn.innerHTML = esc(r[1]);
		if (pc) pc.innerHTML = esc(r[0]);
		window.close();
	}

	// ---- event ----
	var deb = null;
	q.addEventListener('input', function () {
		$('ksClear').style.display = q.value ? 'block' : 'none';
		clearTimeout(deb);
		deb = setTimeout(function () { state.q = q.value; state.start = 0; load(); }, 300);
	});
	$('ksClear').addEventListener('click', function () { q.value = ''; $('ksClear').style.display = 'none'; state.q = ''; state.start = 0; load(); q.focus(); });
	$('ksLen').addEventListener('change', function () { state.length = parseInt(this.value, 10) || 25; state.start = 0; load(); });
	$('ksClose').addEventListener('click', function () { window.close(); });
	pager.addEventListener('click', function (e) { var bt = e.target.closest('button[data-p]'); if (bt && !bt.disabled) goPage(parseInt(bt.getAttribute('data-p'), 10)); });
	body.addEventListener('click', function (e) { var tr = e.target.closest('tr[data-i]'); if (tr) choose(parseInt(tr.getAttribute('data-i'), 10)); });
	body.addEventListener('mousemove', function (e) { var tr = e.target.closest('tr[data-i]'); if (tr) setActive(parseInt(tr.getAttribute('data-i'), 10), false); });
	document.addEventListener('keydown', function (e) {
		var cur = Math.floor(state.start / state.length) + 1;
		if (e.key === 'ArrowDown') { e.preventDefault(); setActive(Math.min(state.active + 1, state.rows.length - 1)); }
		else if (e.key === 'ArrowUp') { e.preventDefault(); setActive(Math.max(state.active - 1, 0)); }
		else if (e.key === 'Enter') { if (state.active >= 0 && !/^(BUTTON|SELECT|A)$/.test((e.target && e.target.tagName) || '')) { e.preventDefault(); choose(state.active); } }
		else if (e.key === 'PageDown') { e.preventDefault(); goPage(cur + 1); }
		else if (e.key === 'PageUp') { e.preventDefault(); goPage(cur - 1); }
		else if (e.key === 'Escape') { if (q.value) { $('ksClear').click(); } else { window.close(); } }
	});

	load();
})();
</script>
</body>
</html>
