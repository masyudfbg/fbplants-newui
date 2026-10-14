/* =====================================================================
   Laporan Management Bulanan (p_lapmanagement_monthly.asp) - tampilan saja.
   - Kartu judul (PT, region, BA, periode) + tombol Export Excel / Cetak
   - Tab Spry bergaya modern; tabel tab laporan keuangan dirapikan
   - Tab Realisasi Budget: judul kolom & Deskripsi tetap terlihat saat digulir,
     warna per periode, baris grup, cari baris, pilih periode, sembunyikan baris nol.
   Data, query dan Export Excel tidak diubah: elemen tambahan dilepas sementara
   saat Export agar isi file Excel sama seperti sebelumnya.
   ===================================================================== */
(function () {
	'use strict';
	var doc = document;
	function txt(el) { return (el.textContent || '').replace(/\s+/g, ' ').trim(); }
	function el(tag, cls, text) {
		var e = doc.createElement(tag);
		if (cls) e.className = cls;
		if (text != null) e.textContent = text;
		return e;
	}
	function param(name) {
		var m = new RegExp('[?&]' + name + '=([^&#]*)', 'i').exec(location.search);
		if (!m) return '';
		try { return decodeURIComponent(m[1].replace(/\+/g, ' ')); } catch (e) { return m[1]; }
	}
	var ZERO = /^-?0([.,]0+)?$/;
	var NEG = /^(-[\d.,]*[1-9][\d.,]*|\([\d.,]+\))$/;
	function markNumber(td) {
		var t = txt(td);
		if (!t) return;
		if (ZERO.test(t)) td.classList.add('lm-zero');
		else if (NEG.test(t)) td.classList.add('lm-neg');
	}

	function init() {
		var outer = doc.getElementById('product-table');
		if (!outer) return;
		doc.body.classList.add('lapmgmt');
		// judul melayang x.js (kelas xthf-site) tidak dipakai: baris judul lama diganti kartu judul
		outer.classList.remove('xthf-site');

		// ---- kerangka luar & kartu judul ----
		var host = null;
		Array.prototype.forEach.call(outer.rows, function (r) {
			if (r.parentNode.parentNode !== outer) return;
			if (r.cells.length === 1 && r.cells[0].tagName === 'TH') r.classList.add('lm-oldtitle');
			else if (!host && r.querySelector('.TabbedPanels')) { host = r.cells[0]; host.classList.add('lm-host'); }
		});

		var company = '', period = '';
		var firstStrong = doc.querySelector('.TabbedPanelsContent table.data strong');
		if (firstStrong) company = txt(firstStrong);
		Array.prototype.some.call(doc.querySelectorAll('.TabbedPanelsContent table.data td'), function (td) {
			var t = txt(td);
			if (/^Per\s+\S+\/\d{4}$/.test(t)) { period = t.replace(/^Per\s+/, ''); return true; }
			return false;
		});
		var region = param('region'), cc = param('company'), ba = param('fcba');
		var hero = el('div', 'lm-hero');
		var main = el('div', 'lm-hero-main');
		main.appendChild(el('div', 'lm-hero-kicker', 'Laporan bulanan'));
		main.appendChild(el('h1', '', 'Laporan Management'));
		var chips = el('div', 'lm-chips');
		function chip(label, value) {
			if (!value) return;
			var c = el('span', 'lm-chip');
			c.appendChild(el('b', '', label));
			c.appendChild(doc.createTextNode(value));
			chips.appendChild(c);
		}
		chip('PT', company || (cc && cc !== '-' ? cc : 'Semua PT'));
		chip('Region', region && region !== '-' ? region : 'Semua region');
		chip('BA', ba && ba !== '-' ? ba : 'Semua BA');
		chip('Periode', period || ((param('bulan') || '') + '/' + (param('tahun') || '')));
		main.appendChild(chips);
		hero.appendChild(main);
		var actions = el('div', 'lm-hero-actions');
		var exp = doc.querySelector('input.tableToExcel');
		if (exp) actions.appendChild(exp);
		var printBtn = el('button', '', 'Cetak');
		printBtn.type = 'button';
		printBtn.addEventListener('click', function () { window.print(); });
		actions.appendChild(printBtn);
		hero.appendChild(actions);
		outer.parentNode.insertBefore(hero, outer);

		// ---- tabel tab laporan keuangan ----
		Array.prototype.forEach.call(doc.querySelectorAll('.TabbedPanelsContent table.data'), function (t) {
			Array.prototype.forEach.call(t.rows, function (r) {
				var cells = Array.prototype.slice.call(r.cells);
				if (!cells.length) return;
				if (cells.every(function (c) { return /border\s*:\s*none/i.test(c.getAttribute('style') || ''); })) { r.classList.add('lm-dochead'); return; }
				if (cells.length > 1 && cells.every(function (c) { return c.classList.contains('ratatengah'); })) { r.classList.add('lm-colhead'); return; }
				var s = cells[0].querySelector('strong, b');
				if (s && txt(s)) r.classList.add('lm-grp');
				cells.forEach(function (c, i) { if (i > 0 && /^[-(]?[\d.,]+\)?$/.test(txt(c))) markNumber(c); });
			});
			accordion(t);
		});

		realisasi();
		patchExport();
	}

	// =====================================================================
	// Accordion baris grup (Neraca, Laba Rugi, Arus Kas, Catatan, Produksi)
	// Tingkat baris dari indentasi teks kolom pertama (spasi/&nbsp; di depan):
	//   ASET (0) > Aset Lancar (3) > KAS DAN SETARA KAS (8). Baris grup = .lm-grp.
	// Menutup grup menyembunyikan baris di bawahnya yang indentasinya lebih dalam.
	// Hanya tampilan: Export Excel & cetak tetap semua baris.
	// =====================================================================
	function indentOf(td) {
		var m = /^[\s\u00a0]*/.exec(td.textContent || '');
		return m ? m[0].length : 0;
	}
	function accordion(t) {
		var rows = Array.prototype.filter.call(t.rows, function (r) {
			return r.cells.length > 1 && !r.classList.contains('lm-dochead') && !r.classList.contains('lm-colhead');
		});
		var groups = rows.filter(function (r) { return r.classList.contains('lm-grp'); });
		if (!groups.length) return;
		// induk tiap baris = baris grup terdekat di atasnya dengan indentasi lebih dangkal
		var stack = [], parent = new Map(), hasChild = new Set();
		rows.forEach(function (r) {
			var lv = indentOf(r.cells[0]);
			while (stack.length && stack[stack.length - 1].lv >= lv) stack.pop();
			var p = stack.length ? stack[stack.length - 1].row : null;
			parent.set(r, p);
			if (p) hasChild.add(p);
			if (r.classList.contains('lm-grp')) stack.push({ row: r, lv: lv });
		});
		var togglers = groups.filter(function (g) { return hasChild.has(g); });
		if (!togglers.length) return;
		function refresh() {
			rows.forEach(function (r) {
				var p = parent.get(r), hide = false;
				while (p) { if (p.classList.contains('lm-closed')) { hide = true; break; } p = parent.get(p); }
				r.classList.toggle('lm-acc-hide', hide);
			});
		}
		togglers.forEach(function (g) {
			g.classList.add('lm-acc');
			var b = el('button', 'lm-acc-btn lm-ui');
			b.type = 'button';
			b.setAttribute('aria-expanded', 'true');
			b.setAttribute('aria-label', 'Buka/tutup ' + txt(g.cells[0]));
			g.cells[0].insertBefore(b, g.cells[0].firstChild);
		});
		function set(g, open) {
			g.classList.toggle('lm-closed', !open);
			var b = g.querySelector('.lm-acc-btn');
			if (b) b.setAttribute('aria-expanded', open ? 'true' : 'false');
		}
		t.addEventListener('click', function (e) {
			var g = e.target.closest('tr.lm-acc');
			if (!g || e.target.closest('a, input, select, textarea')) return;
			set(g, g.classList.contains('lm-closed'));
			refresh();
		});
		// Buka semua / Tutup semua / hanya tingkat teratas
		var bar = el('div', 'lm-acc-bar lm-ui');
		[['Buka semua', function () { togglers.forEach(function (g) { set(g, true); }); }],
		 ['Ringkas', function () { togglers.forEach(function (g) { set(g, !parent.get(g)); }); }],
		 ['Tutup semua', function () { togglers.forEach(function (g) { set(g, false); }); }]].forEach(function (a) {
			var b = el('button', '', a[0]);
			b.type = 'button';
			b.addEventListener('click', function () { a[1](); refresh(); });
			bar.appendChild(b);
		});
		t.parentNode.insertBefore(bar, t);
	}

	// =====================================================================
	// Tab Realisasi Budget
	// =====================================================================
	var rb = null;
	function realisasi() {
		var table = null;
		Array.prototype.some.call(doc.querySelectorAll('.TabbedPanelsContent table'), function (t) {
			if (t.tHead && /Realisasi\s*&\s*Budget/i.test(txt(t.tHead))) { table = t; return true; }
			return false;
		});
		if (!table) return;
		var hr = Array.prototype.slice.call(table.tHead.rows);
		// struktur: [judul] [Deskripsi] [3 periode] [Realisasi/RKB/Var x3] [21 kolom]
		var leaf = hr[hr.length - 1], sub = hr[hr.length - 2], grp = hr[hr.length - 3], descRow = hr[hr.length - 4];
		if (!leaf || leaf.cells.length !== 21 || !grp || grp.cells.length !== 3 || !descRow) return;
		// xTableHeaderFixed (x.js) membuat salinan judul melayang untuk kelas xthf-site;
		// tabel ini sudah punya judul tetap sendiri di wadah gulir.
		table.classList.remove('xthf-site');
		table.classList.add('rb-table');
		var title = hr[0] !== descRow ? hr[0] : null;
		if (title) title.classList.add('rb-oldtitle');
		descRow.cells[0].classList.add('rb-desc');
		grp.classList.add('rb-grp-row');
		sub.classList.add('rb-sub-row');
		leaf.classList.add('rb-leaf');
		var groupNames = [];
		Array.prototype.forEach.call(grp.cells, function (c, g) { c.classList.add('g' + g, 'gs'); groupNames.push(txt(c)); });
		Array.prototype.forEach.call(sub.cells, function (c, k) {
			var g = Math.floor(k / 3), w = k % 3;
			c.classList.add('g' + g);
			if (w === 0) c.classList.add('gs');
		});
		function colClass(c, i) {
			var g = Math.floor(i / 7), w = i % 7;
			c.classList.add('g' + g);
			if (w === 0) c.classList.add('gs');
			if (w >= 3 && w <= 5) c.classList.add('ck');
			if (w === 6) c.classList.add('cv');
		}
		Array.prototype.forEach.call(leaf.cells, colClass);

		var body = table.tBodies[0];
		var rows = [], total = null;
		Array.prototype.forEach.call(body.rows, function (r) {
			if (r.cells.length !== 22) return;
			Array.prototype.forEach.call(r.cells, function (c, i) { if (i > 0) { colClass(c, i - 1); markNumber(c); } });
			if (r.classList.contains('huruftebal') && /^BALANCE$/i.test(txt(r.cells[0]))) { r.classList.add('rb-total'); total = r; return; }
			var first = r.cells[0], s = first.querySelector('strong');
			if (s && txt(s)) r.classList.add('rb-grp');
			else {
				r.classList.add('rb-item');
				// kode pekerjaan diberi badge (teks sel tetap sama)
				var n = s ? s.nextSibling : first.firstChild;
				while (n && n.nodeType !== 3) n = n.nextSibling;
				var m = n && /^(\s*)(\d[\w.\-]*)(\s[\s\S]*)$/.exec(n.nodeValue);
				if (m) {
					var code = el('span', 'rb-code', m[2]);
					n.nodeValue = m[3];
					first.insertBefore(code, n);
				}
			}
			rows.push(r);
		});
		if (!rows.length) return;

		// wadah gulir
		var scroll = el('div', 'rb-scroll');
		table.parentNode.insertBefore(scroll, table);
		scroll.appendChild(table);

		// toolbar
		var bar = el('div', 'rb-bar lm-ui');
		var tl = el('div', 'rb-title');
		tl.appendChild(el('h2', '', 'Realisasi & Budget'));
		var per = '';
		if (title) { var h = title.querySelectorAll('h4, h5'); if (h.length) per = txt(h[h.length - 1]); }
		tl.appendChild(el('p', '', (per ? per + ' \u00b7 ' : '') + 'Realisasi vs RKB per pekerjaan'));
		bar.appendChild(tl);
		var tools = el('div', 'rb-tools');
		var seg = el('div', 'rb-seg');
		seg.setAttribute('role', 'group');
		seg.setAttribute('aria-label', 'Periode yang ditampilkan');
		var dots = ['#4f6f68', '#1B7863', '#12584a'];
		var segBtns = groupNames.map(function (name, g) {
			var b = el('button');
			b.type = 'button';
			b.setAttribute('aria-pressed', 'true');
			var i = el('i'); i.style.background = dots[g];
			b.appendChild(i);
			b.appendChild(doc.createTextNode(name));
			b.addEventListener('click', function () {
				var on = b.getAttribute('aria-pressed') !== 'true';
				// minimal satu periode tetap tampil
				if (!on && segBtns.filter(function (x) { return x.getAttribute('aria-pressed') === 'true'; }).length === 1) return;
				b.setAttribute('aria-pressed', on ? 'true' : 'false');
				table.classList.toggle('rb-hide-' + g, !on);
				sticky();
			});
			seg.appendChild(b);
			return b;
		});
		tools.appendChild(seg);
		var chk = el('label', 'rb-check');
		var cb = el('input'); cb.type = 'checkbox';
		chk.appendChild(cb);
		chk.appendChild(doc.createTextNode('Sembunyikan baris nol'));
		tools.appendChild(chk);
		var sw = el('div', 'rb-search');
		var q = el('input'); q.type = 'search'; q.placeholder = 'Cari kode / nama pekerjaan...';
		q.setAttribute('aria-label', 'Cari baris');
		sw.appendChild(q);
		tools.appendChild(sw);
		bar.appendChild(tools);
		scroll.parentNode.insertBefore(bar, scroll);
		var info = el('div', 'rb-info lm-ui');
		scroll.parentNode.insertBefore(info, scroll.nextSibling);

		var empty = el('tr', 'rb-empty lm-ui');
		var et = el('td', '', 'Tidak ada baris yang cocok.');
		et.colSpan = 22;
		empty.appendChild(et);

		function allZero(r) {
			for (var i = 1; i < r.cells.length; i++) { var t = txt(r.cells[i]); if (t && !ZERO.test(t)) return false; }
			return true;
		}
		function filter() {
			var tokens = q.value.toLowerCase().split(/\s+/).filter(Boolean), shown = 0;
			rows.forEach(function (r) {
				var t = txt(r.cells[0]).toLowerCase();
				var ok = tokens.every(function (k) { return t.indexOf(k) !== -1; }) && !(cb.checked && allZero(r));
				r.classList.toggle('rb-filtered', !ok);
				if (ok) shown++;
			});
			if (!shown) { if (!empty.parentNode) body.insertBefore(empty, total || null); }
			else if (empty.parentNode) empty.parentNode.removeChild(empty);
			info.textContent = (shown === rows.length ? 'Menampilkan ' + rows.length + ' baris' : 'Menampilkan ' + shown + ' dari ' + rows.length + ' baris') +
				' \u00b7 Geser tabel ke samping untuk melihat periode lain';
		}
		q.addEventListener('input', filter);
		q.addEventListener('keydown', function (e) { if (e.key === 'Enter') e.preventDefault(); });
		cb.addEventListener('change', filter);
		filter();

		// judul bertingkat menempel di atas wadah gulir
		function sticky() {
			var vis = hr.filter(function (r) { return r !== title; });
			var base = vis[0].offsetTop;
			vis.forEach(function (r) {
				var top = r.offsetTop - base;
				Array.prototype.forEach.call(r.cells, function (c) { c.style.top = (c.classList.contains('rb-desc') ? 0 : top) + 'px'; });
			});
		}
		rb = { sticky: sticky, bar: bar, info: info, empty: empty };
		sticky();
		var tmr = null;
		function later() { clearTimeout(tmr); tmr = setTimeout(sticky, 80); }
		window.addEventListener('resize', later);
		// tab tersembunyi tingginya 0: hitung ulang saat tab dibuka
		Array.prototype.forEach.call(doc.querySelectorAll('.TabbedPanelsTab'), function (t) { t.addEventListener('click', later); });
	}

	// Export Excel: elemen tambahan dilepas sementara supaya isi file sama seperti sebelumnya
	function patchExport() {
		var orig = window.tableToExcel;
		if (typeof orig !== 'function' || orig.__lm) return;
		var patched = function () {
			var parked = [];
			Array.prototype.forEach.call(doc.querySelectorAll('#product-table .lm-ui'), function (n) {
				parked.push({ n: n, p: n.parentNode, s: n.nextSibling });
				n.parentNode.removeChild(n);
			});
			var codes = [];
			Array.prototype.forEach.call(doc.querySelectorAll('#product-table .rb-code'), function (c) {
				var t = doc.createTextNode(c.textContent);
				c.parentNode.replaceChild(t, c);
				codes.push({ c: c, t: t });
			});
			try { return orig.apply(this, arguments); }
			finally {
				codes.forEach(function (o) { o.t.parentNode.replaceChild(o.c, o.t); });
				parked.reverse().forEach(function (o) { o.p.insertBefore(o.n, o.s); });
			}
		};
		patched.__lm = true;
		window.tableToExcel = patched;
	}

	if (doc.readyState === 'loading') doc.addEventListener('DOMContentLoaded', init);
	else init();
})();
