/* =====================================================================
   Listing BKM Perawatan & Bahan (p_listing_bkm_perawatan.asp) - tampilan saja.
   Kartu judul, ringkasan, tab dengan jumlah data, tabel ala DataTables (FbpTable),
   judul & kolom kunci menempel, total mengikuti pencarian, kolom kosong bisa disembunyikan.
   Export Excel tetap berisi semua baris/kolom dan total asli.
   ===================================================================== */
(function () {
	'use strict';
	var doc = document;
	function txt(el) { return el ? (el.textContent || '').replace(/\s+/g, ' ').trim() : ''; }
	function el(tag, cls, text) { var e = doc.createElement(tag); if (cls) e.className = cls; if (text != null) e.textContent = text; return e; }
	function num(s) { var v = parseFloat(String(s).replace(/,/g, '')); return isNaN(v) ? 0 : v; }
	function fmt(n, dec) { return n.toLocaleString('en-US', { minimumFractionDigits: dec, maximumFractionDigits: dec }); }
	function fmtInt(n) { return n.toLocaleString('id-ID'); }
	function param(name) { var m = new RegExp('[?&]' + name + '=([^&#]*)', 'i').exec(location.search); return m ? decodeURIComponent(m[1].replace(/\+/g, ' ')) : ''; }

	var NUM = /^(no|amount|mandays|over times|conv\. times|unit|bunch weight|kg brondolan|output|rate|qty)$/i;
	var SUM = /^(amount|mandays|over times|conv\. times|unit|bunch weight|kg brondolan|output|rate|qty)$/i;
	var WRAP = /^(employee name|job code|remarks|activity|item bahan)$/i;
	var KEEP = /^(no|no\.bkm|no bkm|date|employee name|item bahan|qty|posting)$/i;
	var EMPTY = /^(|0|0\.0+|-|-\s*-|\/|\/-)$/;
	var seq = 0;

	function setupTable(table, opt) {
		if (!table || !table.tHead || !table.tBodies[0] || !window.FbpTable) return null;
		var heads = Array.prototype.map.call(table.tHead.rows[0].cells, txt);
		var tbody = table.tBodies[0], foot = table.tFoot;
		var rows = Array.prototype.slice.call(tbody.rows);

		// kelas sel: angka rata kanan, teks panjang boleh membungkus, nol dibuat samar, status jadi badge
		heads.forEach(function (h, i) {
			if (NUM.test(h)) table.tHead.rows[0].cells[i].classList.add('num');
		});
		rows.forEach(function (r) {
			Array.prototype.forEach.call(r.cells, function (c, i) {
				var h = heads[i] || '';
				if (NUM.test(h)) { c.classList.add('num'); if (SUM.test(h) && num(txt(c)) === 0) c.classList.add('zero'); }
				if (WRAP.test(h)) c.classList.add('wrap');
				if (/^no\.?\s?bkm$/i.test(h)) c.classList.add('bp-code');
				if (/^posting$/i.test(h) && !c.querySelector('.bp-badge')) {
					var t = txt(c), b = el('span', 'bp-badge' + (/^YA/i.test(t) ? ' ok' : /^TIDAK/i.test(t) ? ' no' : ''), t);
					c.innerHTML = ''; c.appendChild(b);
				}
			});
		});

		// peta sel total -> kolom (berdasarkan colspan)
		var footMap = [];
		if (foot && foot.rows[0]) {
			var acc = 0;
			Array.prototype.forEach.call(foot.rows[0].cells, function (c) {
				var span = c.colSpan || 1, cols = [];
				for (var k = 0; k < span; k++) cols.push(acc + k);
				footMap.push({ cell: c, cols: cols, span: span });
				acc += span;
			});
		}

		// isi asli sel total (untuk Export Excel) dicatat sebelum FbpTable mengubahnya;
		// saat export ditukar sementara tanpa mengganti elemen sel
		var origCells = footMap.map(function (m) { return { html: m.cell.innerHTML, span: m.span }; });

		// wadah: ringkasan, toolbar, area gulir, pager
		var id = 'bp' + (++seq);
		table.setAttribute('data-bp', id);
		var panel = el('div', 'bp-panel');
		var kpis = el('div', 'bp-kpis bp-ui');
		var tools = el('div', 'bp-tools bp-ui');
		var scroll = el('div', 'bp-scroll');
		var pagerBox = el('div', 'bp-foot bp-ui');
		table.parentNode.insertBefore(panel, table);
		panel.appendChild(kpis); panel.appendChild(tools); panel.appendChild(scroll); scroll.appendChild(table); panel.appendChild(pagerBox);

		function updateTotals(hits) {
			footMap.forEach(function (m, idx) {
				if (idx === 0) {
					var lbl = m.cell.querySelector('strong') || m.cell;
					lbl.textContent = 'Total \u00b7 ' + fmtInt(hits.length) + ' baris';
					return;
				}
				if (m.span !== 1 || !SUM.test(heads[m.cols[0]] || '')) return;
				var col = m.cols[0], sum = 0;
				hits.forEach(function (r) { if (r.cells[col]) sum += num(txt(r.cells[col])); });
				var target = m.cell.querySelector('strong') || m.cell;
				target.textContent = fmt(sum, 2);
			});
			if (opt.kpi) opt.kpi(kpis, hits, heads);
		}

		var api = window.FbpTable.init(table, { pageSize: 20, thousands: ',', dateFormat: 'dmy', onChange: updateTotals });
		var tbRow = table.querySelector('tr.fbt-toolbar-row');
		var fbTools = tbRow && tbRow.querySelector('.fbt-tools');
		if (fbTools) tools.appendChild(fbTools);
		if (tbRow) tbRow.parentNode.removeChild(tbRow);
		var fbFoot = panel.parentNode.querySelector('.fbt-foot') || scroll.parentNode.querySelector('.fbt-foot');
		if (fbFoot) pagerBox.appendChild(fbFoot);

		// kolom yang seluruhnya kosong disembunyikan (bisa ditampilkan lagi)
		var emptyCols = [];
		if (opt.hideEmpty && rows.length) {
			heads.forEach(function (h, i) {
				if (KEEP.test(h)) return;
				var all = rows.every(function (r) { return EMPTY.test(txt(r.cells[i])); });
				if (all) emptyCols.push(i);
			});
		}
		if (emptyCols.length) {
			var style = el('style');
			doc.head.appendChild(style);
			var btn = el('button', 'bp-colbtn');
			btn.type = 'button';
			var hidden = true;
			var applyCols = function () {
				style.textContent = hidden ? emptyCols.map(function (i) {
					return 'table[data-bp="' + id + '"] > thead > tr > :nth-child(' + (i + 1) + '), table[data-bp="' + id + '"] > tbody > tr > :nth-child(' + (i + 1) + ')';
				}).join(', ') + ' { display: none; }' : '';
				footMap.forEach(function (m) {
					var vis = m.cols.filter(function (c) { return !hidden || emptyCols.indexOf(c) < 0; }).length;
					m.cell.style.display = vis ? '' : 'none';
					if (vis) m.cell.colSpan = vis;
				});
				btn.textContent = hidden ? emptyCols.length + ' kolom kosong disembunyikan \u00b7 Tampilkan' : 'Sembunyikan ' + emptyCols.length + ' kolom kosong';
				btn.title = emptyCols.map(function (i) { return heads[i]; }).join(', ');
				stickyOffset();
			};
			btn.addEventListener('click', function () { hidden = !hidden; applyCols(); });
			tools.appendChild(btn);
			applyCols();
		}

		// No & No.BKM menempel saat tabel digulir ke samping
		function stickyOffset() {
			if (!opt.sticky) return;
			table.classList.add('bp-sticky');
			var c1 = table.tHead.rows[0].cells[0];
			if (c1 && c1.offsetWidth) table.style.setProperty('--bp-c1', c1.offsetWidth + 'px');
		}
		stickyOffset();
		window.addEventListener('resize', stickyOffset);
		Array.prototype.forEach.call(doc.querySelectorAll('.TabbedPanelsTab'), function (t) { t.addEventListener('click', function () { setTimeout(stickyOffset, 30); }); });

		var saved = null;
		function exportOn() {
			saved = footMap.map(function (m) { return { html: m.cell.innerHTML, span: m.cell.colSpan, disp: m.cell.style.display }; });
			footMap.forEach(function (m, i) { m.cell.innerHTML = origCells[i].html; m.cell.colSpan = origCells[i].span; m.cell.style.display = ''; });
		}
		function exportOff() {
			if (!saved) return;
			footMap.forEach(function (m, i) { m.cell.innerHTML = saved[i].html; m.cell.colSpan = saved[i].span; m.cell.style.display = saved[i].disp; });
			saved = null;
		}
		return { table: table, count: rows.length, api: api, exportOn: exportOn, exportOff: exportOff };
	}

	function kpiCard(label, value, note) {
		return '<div class="bp-kpi"><span>' + label + '</span><b>' + value + '</b>' + (note ? '<small>' + note + '</small>' : '') + '</div>';
	}
	function uniq(hits, col) { var s = {}; hits.forEach(function (r) { var t = txt(r.cells[col]); if (t) s[t] = 1; }); return Object.keys(s).length; }
	function sumCol(hits, col) { var t = 0; hits.forEach(function (r) { t += num(txt(r.cells[col])); }); return t; }
	function idx(heads, re) { for (var i = 0; i < heads.length; i++) if (re.test(heads[i])) return i; return -1; }

	function init() {
		var outer = doc.getElementById('product-table');
		if (!outer || !window.FbpTable) return;
		doc.body.classList.add('bkmlist');

		// kerangka luar: baris judul lama disembunyikan di layar (tetap ikut Export Excel)
		var company = '', ba = '', period = '';
		Array.prototype.forEach.call(outer.tHead ? outer.tHead.rows : [], function (r) {
			if (r.cells.length === 1 && r.cells[0].tagName === 'TH') {
				r.classList.add('bp-oldtitle');
				var h4 = r.querySelector('h4'), h5 = r.querySelector('h5');
				if (h4) company = txt(h4);
				if (h5) ba = txt(h5);
				var t = txt(r);
				if (/^From\s/i.test(t)) period = t.replace(/^From\s+/i, '').replace(/\s+to\s+/i, ' \u2013 ');
			}
		});
		var host = outer.querySelector('.TabbedPanels');
		if (host && host.parentNode && host.parentNode.tagName === 'TD') host.parentNode.classList.add('bp-host');

		var hero = el('div', 'bp-hero');
		var main = el('div', 'bp-hero-main');
		main.appendChild(el('div', 'bp-kicker', 'Listing BKM'));
		main.appendChild(el('h1', '', 'BKM Perawatan & Bahan'));
		var chips = el('div', 'bp-chips');
		function chip(l, v) { if (!v) return; var c = el('span', 'bp-chip'); c.appendChild(el('b', '', l)); c.appendChild(doc.createTextNode(v)); chips.appendChild(c); }
		chip('PT', company.replace(/^PT\.?\s*/i, ''));
		chip('BA', ba);
		// periode dari parameter (mm/dd/yyyy) ditampilkan dd/mm/yyyy; cadangan: teks judul lama
		var dmy = function (v) { var m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(v || ''); return m ? ('0' + m[2]).slice(-2) + '/' + ('0' + m[1]).slice(-2) + '/' + m[3] : v; };
		chip('Periode', param('bkmfrom') ? dmy(param('bkmfrom')) + ' \u2013 ' + dmy(param('bkmto')) : period);
		chip('Dokumen', param('doc'));
		main.appendChild(chips);
		hero.appendChild(main);
		var acts = el('div', 'bp-actions');
		var exp = doc.querySelector('input.tableToExcel');
		if (exp) acts.appendChild(exp);
		var pr = el('button', '', 'Cetak'); pr.type = 'button'; pr.addEventListener('click', function () { window.print(); });
		acts.appendChild(pr);
		hero.appendChild(acts);
		outer.parentNode.insertBefore(hero, outer);

		var upah = setupTable(doc.getElementById('bkm-upah'), {
			hideEmpty: true, sticky: true,
			kpi: function (box, hits, heads) {
				box.innerHTML =
					kpiCard('Baris', fmtInt(hits.length)) +
					kpiCard('BKM', fmtInt(uniq(hits, idx(heads, /^no\.bkm$/i)))) +
					kpiCard('Karyawan', fmtInt(uniq(hits, idx(heads, /^employee name$/i)))) +
					kpiCard('Total HK', fmt(sumCol(hits, idx(heads, /^mandays$/i)), 2), 'Mandays') +
					kpiCard('Total Amount', fmt(sumCol(hits, idx(heads, /^amount$/i)), 2), 'Allowance');
			}
		});
		var bahan = setupTable(doc.getElementById('bkm-bahan'), {
			hideEmpty: false, sticky: false,
			kpi: function (box, hits, heads) {
				box.innerHTML =
					kpiCard('Baris', fmtInt(hits.length)) +
					kpiCard('BKM', fmtInt(uniq(hits, idx(heads, /^no bkm$/i)))) +
					kpiCard('Item Bahan', fmtInt(uniq(hits, idx(heads, /^item bahan$/i)))) +
					kpiCard('Total Qty', fmt(sumCol(hits, idx(heads, /^qty$/i)), 2));
			}
		});

		// jumlah data di tab
		var tabs = doc.querySelectorAll('.TabbedPanelsTab');
		[upah, bahan].forEach(function (t, i) { if (t && tabs[i]) tabs[i].appendChild(el('span', 'bp-count bp-ui', fmtInt(t.count))); });

		// Export Excel: elemen tambahan dilepas & total asli dipulihkan sementara
		var orig = window.tableToExcel;
		if (typeof orig === 'function') {
			window.tableToExcel = function () {
				var parked = [];
				Array.prototype.forEach.call(doc.querySelectorAll('#product-table .bp-ui, #product-table .fbt-empty-row'), function (n) { parked.push({ n: n, p: n.parentNode, s: n.nextSibling }); n.parentNode.removeChild(n); });
				[upah, bahan].forEach(function (t) { if (t) t.exportOn(); });
				try { return orig.apply(this, arguments); }
				finally {
					[upah, bahan].forEach(function (t) { if (t) t.exportOff(); });
					parked.reverse().forEach(function (o) { o.p.insertBefore(o.n, o.s); });
				}
			};
		}
	}

	if (doc.readyState === 'loading') doc.addEventListener('DOMContentLoaded', init); else init();
})();
