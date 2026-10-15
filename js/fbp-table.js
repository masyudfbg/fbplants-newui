/* =====================================================================
   fbp-table.js - tabel laporan ala DataTables (tanpa library):
   cari, jumlah baris per halaman, urut kolom, pagination, total ikut filter.
   Hanya mengatur tampilan baris yang sudah dirender server.
   Pasangan css/fbp-table.css.

   Pemakaian:
     FbpTable.init(document.getElementById('product-table'), {
       pageSize: 10,          // 0 = All
       sumCol: 4,             // indeks kolom yang dijumlah ke td.fbt-sum di tfoot
       countEl: el,           // (opsional) elemen untuk jumlah baris
       onChange: function (rows) {}  // (opsional) dipanggil tiap filter berubah
     });
   Sel judul kolom dengan atribut data-nosort tidak ikut diurutkan
   (mis. kolom checkbox "pilih semua").

   Opsi tambahan (untuk tabel yang diisi AJAX):
       allRows: true      // semua baris tbody dianggap data walau jumlah selnya beda
       thousands: '.'     // angka memakai titik sebagai pemisah ribuan (1.947.006)
       observe: true      // baca ulang baris otomatis saat isi tbody diganti script
   Nilai kembali: { render, rows, refresh }.
   ===================================================================== */
(function (global) {
	function text(el) { return (el.textContent || el.innerText || '').replace(/\s+/g, ' ').trim(); }
	function toNumber(v, thousands) {
		var s = String(v).trim();
		if (thousands === '.') s = s.replace(/\.(?=\d{3}(\D|$))/g, '').replace(/,/g, '.');
		else s = s.replace(/,/g, '');
		return /^-?\d+(\.\d+)?$/.test(s) ? parseFloat(s) : NaN;
	}
	function span(row) {
		var n = 0;
		Array.prototype.forEach.call(row.cells, function (c) { n += c.colSpan || 1; });
		return n;
	}
	function toDate(v) {
		var m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2}))?/.exec(String(v).trim());
		return m ? new Date(+m[3], +m[1] - 1, +m[2], +(m[4] || 0), +(m[5] || 0)).getTime() : NaN;
	}
	// format ribuan dengan 2 desimal (1,234,567.89)
	function fmt(n) { if (Math.abs(n) < 0.005) n = 0; var p = n.toFixed(2).split('.'); return p[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',') + '.' + p[1]; }

	function init(table, opt) {
		if (!table || table.getAttribute('data-fbt')) return null;
		opt = opt || {};
		table.setAttribute('data-fbt', '1');
		var thead = table.tHead, tbody = table.tBodies[0];
		if (!thead || !tbody) return null;
		// Judul bertingkat (rowspan/colspan): opt.leafs = [{cell, col}] judul paling bawah per kolom,
		// opt.colCount = jumlah kolom data, opt.headTop = baris judul teratas (toolbar disisipkan di atasnya)
		var leafs = opt.leafs || null;
		var headRow = opt.headTop || thead.rows[thead.rows.length - 1];
		var nCols = leafs ? opt.colCount : headRow.cells.length, nSpan = leafs ? opt.colCount : span(headRow);
		var rows, others, observer = null;
		function scan() {
			var all = Array.prototype.slice.call(tbody.rows).filter(function (r) { return !r.classList.contains('fbt-empty-row'); });
			// Baris data: jumlah sel = jumlah judul, atau = jumlah kolom bila judul memakai colspan
			rows = all.filter(function (r) {
				return opt.allRows ? r.cells.length > 1 : (r.cells.length === nCols || (nSpan !== nCols && r.cells.length === nSpan && span(r) === nSpan));
			});
			others = all.filter(function (r) { return rows.indexOf(r) === -1; });
		}
		scan();
		function num(v) { return toNumber(v, opt.thousands); }
		function date(v) {
			if (opt.dateFormat !== 'dmy') return toDate(v);
			var m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(String(v).trim());
			return m ? new Date(+m[3], +m[2] - 1, +m[1]).getTime() : NaN;
		}
		var state = { page: 1, size: opt.pageSize === undefined ? 10 : opt.pageSize, q: '', col: -1, dir: 1 };

		// ---- toolbar (baris di thead, tepat di atas judul kolom) ----
		var tr = document.createElement('tr');
		tr.className = 'fbt-toolbar-row';
		var th = document.createElement('th');
		th.colSpan = nSpan;
		th.innerHTML =
			'<div class="fbt-tools">' +
			'<label>Tampilkan <select class="fbt-size" aria-label="Jumlah baris per halaman">' +
			'<option value="10">10</option><option value="20">20</option><option value="50">50</option><option value="100">100</option><option value="0">All</option>' +
			'</select> data</label>' +
			'<div class="fbt-search"><input type="search" placeholder="Cari di tabel..." aria-label="Cari di tabel" autocomplete="off"></div>' +
			'</div>';
		tr.appendChild(th);
		thead.insertBefore(tr, headRow);
		var sizeSel = th.querySelector('.fbt-size');
		var input = th.querySelector('input');
		sizeSel.value = String(state.size);

		// ---- footer (info + pager) setelah tabel ----
		var foot = document.createElement('div');
		foot.className = 'fbt-foot';
		foot.innerHTML = '<div class="fbt-info" aria-live="polite"></div><div class="fbt-pager" role="navigation" aria-label="Halaman tabel"></div>';
		table.parentNode.insertBefore(foot, table.nextSibling);
		var info = foot.querySelector('.fbt-info'), pager = foot.querySelector('.fbt-pager');

		var empty = document.createElement('tr');
		empty.className = 'fbt-empty-row';
		empty.innerHTML = '<td colspan="' + nSpan + '">Tidak ada data yang cocok dengan pencarian.</td>';

		var sumCell = opt.sumCol !== undefined && table.tFoot ? table.tFoot.querySelector('.fbt-sum') : null;
		var countCell = table.tFoot ? table.tFoot.querySelector('.fbt-count') : null;

		// ---- urut kolom ----
		// posisi kolom data untuk tiap judul (judul ber-colspan menggeser kolom sesudahnya)
		var colOf = [], acc = 0;
		Array.prototype.forEach.call(headRow.cells, function (c) { colOf.push(acc); acc += c.colSpan || 1; });
		var sortList = leafs
			? leafs.map(function (l, i) { return { cell: l.cell, col: l.col, idx: i }; })
			: Array.prototype.map.call(headRow.cells, function (c, i) { return { cell: c, col: colOf[i], idx: i }; });
		function cellAt(r, idx) {
			var s = sortList[idx];
			return leafs ? r.cells[s.col] : r.cells[r.cells.length === nSpan ? s.col : s.idx];
		}

		sortList.forEach(function (sl, idx) {
			var cell = sl.cell;
			// Kolom berisi kontrol (mis. checkbox pilih semua) tidak diurutkan: <td data-nosort>.
			// Judul ber-colspan (mis. beberapa kolom ikon aksi) juga tidak diurutkan.
			if (cell.hasAttribute('data-nosort') || (cell.colSpan || 1) > 1) return;
			cell.classList.add('fbt-sortable');
			cell.setAttribute('tabindex', '0');
			cell.setAttribute('title', 'Klik untuk mengurutkan');
			function sort(e) {
				if (e.type === 'keydown' && e.key !== 'Enter' && e.key !== ' ') return;
				e.preventDefault();
				state.dir = state.col === idx ? -state.dir : 1;
				state.col = idx;
				sortList.forEach(function (o, i) {
					if (i === idx) o.cell.setAttribute('aria-sort', state.dir === 1 ? 'ascending' : 'descending');
					else o.cell.removeAttribute('aria-sort');
				});
				sortRows();
				state.page = 1;
				render();
			}
			cell.addEventListener('click', sort);
			cell.addEventListener('keydown', sort);
		});

		function sortRows() {
			var idx = state.col;
			if (idx < 0) return;
			rows.sort(function (a, b) {
				var ca = cellAt(a, idx), cb = cellAt(b, idx);
				var x = ca ? text(ca) : '', y = cb ? text(cb) : '';
				var nx = num(x), ny = num(y);
				if (!isNaN(nx) && !isNaN(ny)) return (nx - ny) * state.dir;
				var dx = date(x), dy = date(y);
				if (!isNaN(dx) && !isNaN(dy)) return (dx - dy) * state.dir;
				return x.localeCompare(y, 'id', { numeric: true, sensitivity: 'base' }) * state.dir;
			});
			rows.forEach(function (r) { tbody.appendChild(r); });
			if (observer) observer.takeRecords();   // urutan diubah tabel sendiri, bukan data baru
		}

		function button(label, page, o) {
			o = o || {};
			return '<button type="button" data-page="' + page + '"' + (o.cur ? ' aria-current="page"' : '') +
				(o.dis ? ' disabled' : '') + (o.aria ? ' aria-label="' + o.aria + '"' : '') + '>' + label + '</button>';
		}

		function render() {
			var tokens = state.q.toLowerCase().split(/\s+/).filter(Boolean);
			var hits = [];
			rows.forEach(function (r) {
				var t = text(r).toLowerCase();
				var ok = tokens.every(function (k) { return t.indexOf(k) !== -1; });
				r.classList.toggle('fbt-filter', !ok);
				r.classList.add('fbt-page');
				if (ok) hits.push(r);
			});
			var size = state.size > 0 ? state.size : Math.max(hits.length, 1);
			var pages = Math.max(1, Math.ceil(hits.length / size));
			state.page = Math.min(Math.max(state.page, 1), pages);
			var start = (state.page - 1) * size, end = Math.min(start + size, hits.length);
			for (var i = start; i < end; i++) hits[i].classList.remove('fbt-page');

			if (rows.length && !hits.length) { if (!empty.parentNode) tbody.appendChild(empty); }
			else if (empty.parentNode) empty.parentNode.removeChild(empty);
			if (observer) observer.takeRecords();   // baris "tidak ada data" bukan data baru
			others.forEach(function (r) { r.style.display = rows.length ? 'none' : ''; });

			if (!hits.length) info.innerHTML = rows.length ? 'Tidak ada data yang ditampilkan' : 'Tidak ada data';
			else if (state.size === 0) info.innerHTML = 'Menampilkan semua <b>' + hits.length + '</b> data';
			else info.innerHTML = 'Menampilkan <b>' + (start + 1) + '&ndash;' + end + '</b> dari <b>' + hits.length + '</b> data';
			if (tokens.length && rows.length) info.innerHTML += ' (difilter dari ' + rows.length + ' total)';

			var html = button('&lsaquo;', state.page - 1, { dis: state.page <= 1, aria: 'Halaman sebelumnya' });
			var list = [];
			if (pages <= 7) { for (var p = 1; p <= pages; p++) list.push(p); }
			else {
				var from = Math.max(2, state.page - 1), to = Math.min(pages - 1, state.page + 1);
				if (state.page <= 3) { from = 2; to = 4; }
				if (state.page >= pages - 2) { from = pages - 3; to = pages - 1; }
				list.push(1); if (from > 2) list.push(0);
				for (var q = from; q <= to; q++) list.push(q);
				if (to < pages - 1) list.push(0); list.push(pages);
			}
			list.forEach(function (n) { html += n === 0 ? '<span>&hellip;</span>' : button(String(n), n, { cur: n === state.page, aria: 'Halaman ' + n }); });
			html += button('&rsaquo;', state.page + 1, { dis: state.page >= pages, aria: 'Halaman berikutnya' });
			pager.innerHTML = rows.length ? html : '';

			if (sumCell) {
				var sum = 0;
				hits.forEach(function (r) { var n = r.cells[opt.sumCol] ? num(text(r.cells[opt.sumCol])) : NaN; if (!isNaN(n)) sum += n; });
				sumCell.textContent = fmt(sum);
			}
			if (countCell) countCell.textContent = hits.length;
			if (typeof opt.onChange === 'function') opt.onChange(hits);
		}

		sizeSel.addEventListener('change', function () { state.size = parseInt(this.value, 10) || 0; state.page = 1; render(); });
		input.addEventListener('input', function () { state.q = this.value; state.page = 1; render(); });
		input.addEventListener('keydown', function (e) { if (e.key === 'Enter') e.preventDefault(); });
		pager.addEventListener('click', function (e) {
			var b = e.target.closest ? e.target.closest('button') : null;
			if (!b || b.disabled) return;
			state.page = parseInt(b.getAttribute('data-page'), 10);
			render();
		});

		// Baca ulang baris (mis. tbody diisi ulang lewat AJAX); urutan & kata pencarian dipertahankan
		function refresh() {
			scan();
			sortRows();
			state.page = 1;
			render();
		}
		if (opt.observe && global.MutationObserver) {
			observer = new MutationObserver(function () { refresh(); });
			observer.observe(tbody, { childList: true });
		}

		render();
		return { render: render, rows: function () { return rows.slice(); }, refresh: refresh };
	}

	// Export Excel (print_excel.asp): ekspor hasil filter lengkap, tanpa toolbar & tanpa pagination
	function patchExport() {
		var orig = global.tableToExcel;
		if (typeof orig !== 'function' || orig.fbtPatched) return;
		var patched = function (tableID, filename) {
			var src = document.getElementById(tableID);
			if (!src || !src.getAttribute('data-fbt')) return orig.apply(this, arguments);
			var clone = src.cloneNode(true);
			Array.prototype.forEach.call(clone.querySelectorAll('.fbt-toolbar-row, .fbt-filter, .fbt-empty-row, .fbt-rdetail'), function (el) { el.parentNode.removeChild(el); });
			Array.prototype.forEach.call(clone.querySelectorAll('.fbt-page'), function (el) { el.classList.remove('fbt-page'); });
			clone.id = tableID + '-export';
			clone.style.display = 'none';
			document.body.appendChild(clone);
			try { orig.call(this, clone.id, filename); } finally { document.body.removeChild(clone); }
		};
		patched.fbtPatched = true;
		global.tableToExcel = patched;
	}

	global.FbpTable = { init: init, patchExport: patchExport };
})(window);
