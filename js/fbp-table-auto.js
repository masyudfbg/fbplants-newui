/* =====================================================================
   fbp-table-auto.js - tampilan laporan modern + tabel ala DataTables
   (js/fbp-table.js) secara otomatis untuk tabel laporan #product-table.
   Dimuat dari Bottom_Report.asp sehingga berlaku di semua laporan yang memakainya.

   Hanya tampilan: baris tetap dari server, tidak ada request baru.

   1) Tabel ala DataTables (cari, jumlah baris, urut, pagination). Tabel DILEWATI bila
        - sudah diaktifkan sendiri oleh halaman (data-fbt), atau diberi data-fbt-auto="off"
        - tidak punya <thead>, atau judul kolom bertingkat (rowspan)
        - ada baris non-data (subtotal / judul grup) di tengah data
      Baris non-data di ujung tabel (mis. baris Total) dipindah ke <tfoot> agar tetap tampil.
      Header lengket memakai CSS sticky; salinan header xTableHeaderFixed disembunyikan
      karena lebar kolomnya tidak ikut berubah saat pindah halaman / pencarian.

   2) Kulit laporan (seperti laporan Purchase Request): kop kartu dari #header-wrapper
      (PT, BA, judul) + label filter dari baris "Selected by", tombol Cetak & Export Excel
      di kop, kartu ringkasan, tabel di dalam panel, badge status.
      Tidak dipasang pada laporan yang sudah punya tampilan sendiri (tabel data-fbt dari
      halaman) atau yang diberi <body data-rpt-skin="off">.
      Baris kop lama di dalam <thead> disembunyikan di layar tetapi tetap ikut Export Excel.
   ===================================================================== */
(function (global) {
	function cells(row) { return row ? row.cells.length : 0; }
	function spanOf(row) {
		var n = 0;
		Array.prototype.forEach.call(row.cells, function (c) { n += c.colSpan || 1; });
		return n;
	}
	function hasRowspanInto(thead, last) {
		// sel rowspan di baris judul sebelumnya yang "turun" ke baris judul terakhir
		var rows = thead.rows, li = Array.prototype.indexOf.call(rows, last);
		for (var i = 0; i < li; i++) {
			for (var j = 0; j < rows[i].cells.length; j++) {
				if ((rows[i].cells[j].rowSpan || 1) > li - i) return true;
			}
		}
		return false;
	}
	function text(el) { return el ? (el.textContent || '').replace(/\s+/g, ' ').trim() : ''; }
	function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
	function toNum(s) {
		s = String(s).replace(/,/g, '').trim();
		return /^-?\d+(\.\d+)?$/.test(s) ? parseFloat(s) : NaN;
	}
	function fmt(n) {
		if (Math.abs(n) < 0.005) n = 0;   // sisa pembulatan desimal (mis. jurnal seimbang) bukan "-0"
		var d = Math.abs(n % 1) > 0.0001 ? 2 : 0;
		return n.toFixed(d).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
	}
	// Judul kapital -> Title Case; kata pendek (PO, PR, BKM, TBS) tetap kapital
	function titleCase(s) {
		if (s !== s.toUpperCase()) return s;
		return s.toLowerCase().replace(/[a-z0-9.&]+/g, function (w) {
			return w.length <= 3 && /^[a-z]+$/.test(w) && !/^(dan|and|of|per|ke|di|by|to|the)$/.test(w) ? w.toUpperCase() : w.charAt(0).toUpperCase() + w.slice(1);
		});
	}
	function stateClass(s) {
		s = String(s).toLowerCase();
		if (/cancel|reject|batal|void|tidak/.test(s)) return 'bad';
		if (/approv|post|order|close|complete|release|selesai|^ya$|^y$/.test(s)) return 'ok';
		if (/progress|authoriz|confirm|review/.test(s)) return 'info';
		if (/plan|draft|new|open|wait|pending/.test(s)) return 'warn';
		return '';
	}

	// Label judul dari nama kolom database (dipakai bila judul kolom kurang)
	var FIELD_LABELS = {
		FCENTRY: 'User Input', FCEDIT: 'User Edit', LASTUPDATE: 'Last Update', INPUTDATE: 'Tgl Input',
		DATECREATED: 'Date Created', FCBA: 'BA', FCIP: 'IP', IPADDRESS: 'IP Address', FCCODE: 'Code',
		FCNAME: 'Name', ACTIVATION: 'Active', ROWSTATE: 'Status', ROWSTATEID: 'Status', REMARKS: 'Remarks'
	};
	function fieldLabel(name, idx) {
		if (!name) return 'Kolom ' + (idx + 1);
		var up = String(name).toUpperCase();
		if (FIELD_LABELS[up]) return FIELD_LABELS[up];
		return String(name).toLowerCase().replace(/_/g, ' ').replace(/\b[a-z]/g, function (c) { return c.toUpperCase(); });
	}

	// Alasan tabel dilewati dicatat di atribut data-fbt-skip (untuk pemeriksaan)
	function skip(table, why) { table.setAttribute('data-fbt-skip', why); return false; }

	// Judul kolom bertingkat -> peta kolom. Baris judul kolom = baris thead yang punya lebih dari
	// satu sel (baris kop / "Selected by" hanya satu sel). Hasil: kolom data, judul paling bawah
	// per kolom (leafs, untuk urut), dan baris judul teratas.
	function headerGrid(thead) {
		var hr = Array.prototype.filter.call(thead.rows, function (r) { return cells(r) > 1; });
		if (!hr.length) return null;
		var grid = [];
		hr.forEach(function (r, ri) {
			grid[ri] = grid[ri] || [];
			var c = 0;
			Array.prototype.forEach.call(r.cells, function (cell) {
				while (grid[ri][c]) c++;
				var rs = Math.min(cell.rowSpan || 1, hr.length - ri), cs = cell.colSpan || 1;
				for (var dr = 0; dr < rs; dr++) {
					grid[ri + dr] = grid[ri + dr] || [];
					for (var dc = 0; dc < cs; dc++) grid[ri + dr][c + dc] = cell;
				}
				c += cs;
			});
		});
		var last = grid[grid.length - 1], colCount = last.length, leafs = [];
		for (var j = 0; j < colCount; j++) {
			var cell = last[j];
			if (!cell) return null;                     // lubang di peta judul: struktur tidak dikenal
			if ((cell.colSpan || 1) === 1 && (j === 0 || last[j - 1] !== cell)) leafs.push({ cell: cell, col: j });
		}
		return { colCount: colCount, leafs: leafs, top: hr[0], rows: hr };
	}

	function headRowOf(table) {
		var thead = table.tHead;
		if (!thead) return null;
		for (var i = thead.rows.length - 1; i >= 0; i--) { if (cells(thead.rows[i])) return thead.rows[i]; }
		return null;
	}

	function enhance(table, onChange) {
		if (!table || table.getAttribute('data-fbt') || table.getAttribute('data-fbt-auto') === 'off') return false;
		var thead = table.tHead, tbody = table.tBodies[0];
		if (!thead || !tbody) return skip(table, 'struktur');

		// baris judul kolom = baris thead terakhir yang berisi sel
		var head = headRowOf(table);
		if (!head) return skip(table, 'tanpa-judul');
		var n = cells(head);
		var nSpan = spanOf(head);
		if (n < 1) return skip(table, 'tanpa-judul');
		// Judul bertingkat (rowspan/colspan, mis. "Sidik Jari: Check In | Check Out")
		var multi = null;
		if (hasRowspanInto(thead, head)) {
			multi = headerGrid(thead);
			if (!multi) return skip(table, 'judul-rowspan');
			n = nSpan = multi.colCount;
		}
		if (!multi && n < 2) return skip(table, 'judul-rowspan');
		// Judul kolom lebih sedikit dari isi baris (mis. PrintData.asp mencetak semua kolom query
		// tetapi judulnya daftar tetap): lengkapi judul agar selebar tabel
		if (!multi && n === nSpan) {
			var counts = Array.prototype.filter.call(tbody.rows, function (r) { return cells(r) > 1; }).map(cells);
			var m = counts.length ? counts[0] : 0;
			if (m > n && counts.every(function (c) { return c === m; })) {
				var names = global.FBT_FIELD_NAMES || [];
				for (var x = n; x < m; x++) {
					var th = document.createElement('th');
					th.textContent = fieldLabel(names[x], x);
					head.appendChild(th);
				}
				n = nSpan = m;
			}
		}
		// <tbody> tambahan: baris TOTAL yang ditulis setelah </tbody> membuat browser membuat tbody kedua
		// (mis. p_salesshipment.asp). Bila tbody tambahan hanya berisi baris non-data, gabungkan ke tbody
		// utama (nanti dipindah ke tfoot sebagai baris total). Tabel yang dikelompokkan per tbody dilewati.
		if (table.tBodies.length > 1) {
			var extra = Array.prototype.slice.call(table.tBodies, 1);
			var onlyTotals = extra.every(function (tb) {
				return Array.prototype.every.call(tb.rows, function (r) { return !(cells(r) === nSpan && spanOf(r) === nSpan); });
			});
			if (!onlyTotals) return skip(table, 'struktur');
			extra.forEach(function (tb) { while (tb.rows.length) tbody.appendChild(tb.rows[0]); tb.parentNode.removeChild(tb); });
		}
		// jangan pindahkan baris judul kosong ke posisi terakhir: FbpTable memakai baris thead terakhir
		if (thead.rows[thead.rows.length - 1] !== head) {
			for (var k = thead.rows.length - 1; thead.rows[k] !== head; k--) {
				if (cells(thead.rows[k])) return skip(table, 'judul-ganda');
				thead.rows[k].parentNode.removeChild(thead.rows[k]);
			}
		}

		var rows = Array.prototype.slice.call(tbody.rows);
		// baris data: satu sel per kolom (judul ber-colspan, mis. kolom ikon aksi, tetap didukung)
		var isData = rows.map(function (r) { return cells(r) === nSpan && spanOf(r) === nSpan; });
		var firstData = isData.indexOf(true), lastData = isData.lastIndexOf(true);
		if (firstData === -1) {
			// Tidak ada data sama sekali: tampilkan pesan, bukan judul kolom tanpa isi
			var textRows = rows.filter(function (r) { return (r.textContent || '').trim(); });
			// satu baris pesan bawaan laporan (mis. "No Data Available") = pesan kosong
			var msgRow = textRows.length === 1 && cells(textRows[0]) === 1 ? textRows[0] : null;
			// hanya baris Total (bernilai 0) tanpa data: pindah ke tfoot, lalu tampilkan pesan kosong
			var isTotal = function (r) { return r.hasAttribute('bgcolor') || /^t\s*o\s*t\s*a\s*l/i.test(text(r)); };
			if (textRows.length && !msgRow && textRows.every(isTotal)) {
				var tf = table.tFoot || table.createTFoot();
				textRows.forEach(function (r) { tf.appendChild(r); });
				textRows = [];
			}
			if (textRows.length && !msgRow) return skip(table, 'tanpa-baris-data');
			// tetap tampil ala DataTables: toolbar + pesan kosong
			if (msgRow) {
				msgRow.className += ' fbt-auto-empty';
			} else {
				var er = tbody.insertRow(-1);
				er.className = 'fbt-auto-empty';
				var ec = er.insertCell(0);
				ec.colSpan = nSpan;
				ec.textContent = 'Tidak ada data untuk filter ini.';
			}
		} else {
			for (var a = 0; a < firstData; a++) { if ((rows[a].textContent || '').trim()) return skip(table, 'grup-awal'); }   // judul grup di awal
			for (var b = firstData; b <= lastData; b++) { if (!isData[b]) return skip(table, 'subtotal-tengah'); }            // subtotal di tengah
			var trailing = rows.slice(lastData + 1).filter(function (r) { return (r.textContent || '').trim() || r.querySelector('input,select'); });

			// baris Total di ujung -> tfoot (tetap tampil, tidak ikut pagination)
			if (trailing.length) {
				var tfoot = table.tFoot || table.createTFoot();
				trailing.forEach(function (r) { tfoot.appendChild(r); });
			}
		}

		table.setAttribute('data-fbt-auto', '1');
		if (table.xthfHdrTblId) {
			var clone = document.getElementById(table.xthfHdrTblId);
			if (clone) clone.className += ' fbt-xthf-off';
		}
		global.FbpTable.patchExport();
		var opts = { pageSize: 20, onChange: onChange };
		if (multi) {
			opts.leafs = multi.leafs; opts.colCount = multi.colCount; opts.headTop = multi.top;
			table.__fbtLeafs = multi.leafs;
			table.className += ' fbt-multi';
		}
		global.FbpTable.init(table, opts);
		return true;
	}

	// ---------------------------------------------------------------------
	// Kulit laporan
	// ---------------------------------------------------------------------
	// Kop laporan: #header-wrapper berjudul (h2/h3/h4). Di beberapa halaman REPORT ada
	// #header-wrapper lain berisi form filter; yang itu tidak dipakai dan tidak disembunyikan.
	function pickHeaderWrapper(table) {
		var all = document.querySelectorAll('[id="header-wrapper"]'), best = null;
		Array.prototype.forEach.call(all, function (el) {
			if (best && table.tHead && table.tHead.contains(best)) return;
			if (el.querySelector('form, select, input, textarea')) return;
			if (!el.querySelector('h1, h2, h3, h4')) return;
			best = el;
		});
		return best;
	}

	function buildHead(table) {
		var hw = pickHeaderWrapper(table);
		var company = '', ba = '', title = '', notes = [], select = '';
		if (hw) {
			company = text(hw.querySelector('h4'));
			if (!company.replace(/^PT\.?\s*/i, '')) company = '';   // nama perusahaan kosong: jangan tampilkan 'PT.' saja
			ba = text(hw.querySelector('h5'));
			var h = hw.querySelector('h2') || hw.querySelector('h3') || hw.querySelector('h1');
			title = text(h);
			// keterangan lain di bawah judul (mis. catatan laporan)
			var sub = hw.querySelector('#subheader');
			if (sub) Array.prototype.forEach.call(sub.children, function (c) { if (c !== h && text(c)) notes.push(text(c)); });
		}
		if (!title) title = text(document.querySelector('title')).replace(/^.*?reporting\s*-\s*/i, '');
		if (!title) return null;

		// baris kop di dalam thead: disembunyikan di layar, tetap ikut Export Excel
		var thead = table.tHead, head = headRowOf(table);
		if (thead) {
			Array.prototype.forEach.call(thead.rows, function (r) {
				// rpt-keep: baris grup judul kolom yang tetap ditampilkan (mis. Biodata / Pekerjaan / Pajak)
				if (r === head || r.className.indexOf('fbt-toolbar-row') > -1 || r.className.indexOf('rpt-keep') > -1) return;
				var t = text(r);
				// baris judul kolom (lebih dari satu sel, bukan kop) tetap tampil, mis. judul bertingkat
				if (cells(r) > 1 && !r.querySelector('[id="header-wrapper"]') && !/^selected\s*by/i.test(t)) return;
				if (/^selected\s*by/i.test(t)) select = t.replace(/^selected\s*by\s*:?\s*/i, '');
				r.className += ' rpt-xls';
			});
		}
		if (hw && !(thead && thead.contains(hw))) hw.className += ' rpt-orig-hidden';

		var card = document.createElement('div');
		card.className = 'rpt-head';
		var chip = select && !/^all data$/i.test(select)
			? '<span class="rpt-chip"><b>Filter</b>' + esc(select) + '</span>'
			: '<span class="rpt-chip"><b>Filter</b>Semua data</span>';
		card.innerHTML =
			'<div>' +
			(company ? '<div class="rpt-company">' + esc(company) + '</div>' : '') +
			(ba ? '<div class="rpt-ba">' + esc(ba) + '</div>' : '') +
			'<h2 class="rpt-title">' + esc(titleCase(title)) + '</h2>' +
			(notes.length ? '<div class="rpt-sub">' + esc(notes.join(' · ')) + '</div>' : '') +
			'<div class="rpt-chips">' + chip + '</div>' +
			'</div>' +
			'<div class="rpt-actions"><button type="button" class="rpt-print" onclick="window.print()">Cetak</button></div>';
		// tombol Export Excel (print_excel.asp) dipindah ke kop; laporan tanpa tombol itu diberi yang sama
		var xl = document.querySelector('input[onclick*="tableToExcel"]');
		if (!xl && table.id) {
			if (typeof global.tableToExcel !== 'function') {
				// sama dengan print_excel.asp
				global.tableToExcel = function (tableID, filename) {
					var t = document.getElementById(tableID);
					var html = '<html><head><meta charset="UTF-8"></head><body>' + t.outerHTML + '</body></html>';
					var url = URL.createObjectURL(new Blob([html], { type: 'application/vnd.ms-excel' }));
					var a = document.createElement('a');
					a.href = url; a.download = filename || 'Export.xls';
					document.body.appendChild(a); a.click(); document.body.removeChild(a);
					URL.revokeObjectURL(url);
				};
			}
			xl = document.createElement('input');
			xl.type = 'button';
			xl.value = 'Export to Excel';
			xl.setAttribute('onclick', "tableToExcel('" + table.id + "', 'Report.xls')");
			xl.style.cssText = 'padding: 10px; color: white; background-color: green; border-radius: 5px';
		}
		if (xl) card.querySelector('.rpt-actions').appendChild(xl);
		global.FbpTable.patchExport();   // ekspor tetap semua baris hasil pencarian, tanpa toolbar
		return card;
	}

	// kolom status & kolom nominal utama untuk kartu ringkasan
	function pickColumns(table, rows) {
		var head = headRowOf(table), res = { status: -1, money: -1, moneyLabel: '' };
		if (!head) return res;
		// daftar {cell, col}: dari peta judul bertingkat, atau dari baris judul terakhir
		var list = table.__fbtLeafs;
		if (!list) {
			var acc = 0;
			list = Array.prototype.map.call(head.cells, function (c) { var o = { cell: c, col: acc }; acc += c.colSpan || 1; return o; });
		}
		var sample = rows.slice(0, 30);
		list.forEach(function (o) {
			var c = o.cell, t = text(c), col = o.col;
			if ((c.colSpan || 1) > 1) return;
			// judul persis Status / Rowstate / Posting (bukan "State" provinsi atau "Status Perkawinan")
			if (res.status < 0 && /^(status|rowstate|row state|posting|status posting|status dokumen)$/i.test(t)) res.status = col;
			if (/amount|total|nilai|grand|jumlah|nominal|harga total/i.test(t) && !/qty|quantity|%|persen/i.test(t)) {
				var ok = sample.length && sample.every(function (r) { var x = text(r.cells[col]); return x === '' || !isNaN(toNum(x)); });
				if (ok) { res.money = col; res.moneyLabel = t; }   // kolom paling kanan yang cocok
			}
		});
		return res;
	}

	// Sel tidak terlipat secara default; kolom yang isinya panjang (dilihat dari contoh baris)
	// boleh terlipat dengan lebar minimal agar baris tidak menjadi sangat tinggi
	function layoutColumns(table, rows) {
		var n = 0, sample = rows.slice(0, 60), longCols = [];
		sample.forEach(function (r) { n = Math.max(n, r.cells.length); });
		for (var c = 0; c < n; c++) {
			var total = 0, cnt = 0;
			sample.forEach(function (r) { if (r.cells[c]) { total += text(r.cells[c]).length; cnt++; } });
			if (cnt && total / cnt > 28) longCols.push(c + 1);
		}
		var css = 'body.rpt-skin #product-table > tbody > tr > td { white-space: nowrap; }';
		longCols.forEach(function (k) {
			css += ' body.rpt-skin #product-table > tbody > tr > td:nth-child(' + k + ') { white-space: normal; min-width: 260px; max-width: 420px; }';
		});
		var st = document.createElement('style');
		st.appendChild(document.createTextNode(css));
		document.head.appendChild(st);
	}

	function badgeStatus(rows, col) {
		if (col < 0) return;
		rows.forEach(function (r) {
			var td = r.cells[col];
			if (!td || td.querySelector('.rpt-badge, img, a, input, select')) return;
			var t = text(td);
			if (!t || t.length > 30) return;
			td.innerHTML = '<span class="rpt-badge ' + stateClass(t) + '">' + esc(t) + '</span>';
		});
	}

	function skin(table, wasEnhanced) {
		var body = document.body;
		if (!body || body.getAttribute('data-rpt-skin') === 'off' || body.className.indexOf('rpt-skin') > -1) return;
		var card = buildHead(table);
		if (!card) return;
		body.className += ' rpt-skin';

		var panel = document.createElement('div');
		panel.className = 'rpt-panel';
		var tools = document.createElement('div');
		tools.className = 'rpt-tools';
		var scroll = document.createElement('div');
		scroll.className = 'rpt-scroll rpt-scroll-fixed';
		// salinan header xTableHeaderFixed tidak dipakai lagi (header lengket di dalam panel)
		if (table.xthfHdrTblId) {
			var clone = document.getElementById(table.xthfHdrTblId);
			if (clone && clone.className.indexOf('fbt-xthf-off') < 0) clone.className += ' fbt-xthf-off';
		}
		Array.prototype.forEach.call(document.querySelectorAll('table.xthf-no-print'), function (c) {
			if (c.className.indexOf('fbt-xthf-off') < 0) c.className += ' fbt-xthf-off';
		});
		table.parentNode.insertBefore(card, table);
		var summary = document.createElement('div');
		summary.className = 'fbt-summary rpt-summary';
		table.parentNode.insertBefore(summary, table);
		table.parentNode.insertBefore(panel, table);
		panel.appendChild(tools);
		panel.appendChild(scroll);
		scroll.appendChild(table);
		freezeHead(table);
		// toolbar & pager FbpTable di luar area gulir
		var tb = table.querySelector('.fbt-toolbar-row .fbt-tools');
		if (tb) tools.appendChild(tb);
		var foot = document.querySelector('.fbt-foot');
		if (foot && wasEnhanced) panel.appendChild(foot);
		return summary;
	}

	// Freeze header bertingkat di area gulir laporan (mis. REPORT/p_vra.asp: 3 baris judul + rowspan).
	// Dulu hanya baris judul terakhir yang lengket, sehingga potongan judul hijau "nyangkut" di atas data
	// sementara baris judul di atasnya ikut tergulir. Sekarang semua baris judul yang tampil lengket,
	// masing-masing dengan top = jarak dari baris judul pertama. Tabel .fbt-multi sudah lengket lewat <thead>.
	function freezeHead(table) {
		var thead = table.tHead;
		if (!thead) return;
		function apply() {
			if (table.className.indexOf('fbt-multi') > -1) return;
			var rows = Array.prototype.filter.call(thead.rows, function (r) {
				return r.offsetHeight > 0 && r.className.indexOf('fbt-toolbar-row') < 0;
			});
			if (rows.length < 2) return;
			var base = rows[0].offsetTop;
			rows.forEach(function (r, i) {
				var top = r.offsetTop - base;
				Array.prototype.forEach.call(r.cells, function (c) {
					c.classList.add('fbt-freeze');
					c.style.top = top + 'px';
					c.style.zIndex = String(10 - Math.min(i, 5));
				});
			});
		}
		apply();
		var t = null;
		function later() { clearTimeout(t); t = setTimeout(apply, 120); }
		global.addEventListener('resize', later);
		if (global.ResizeObserver) new ResizeObserver(later).observe(thead);
		if (document.fonts && document.fonts.ready) document.fonts.ready.then(later);
	}

	function renderSummary(summary, table, rows, cols) {
		if (!summary) return;
		function card(label, value, note, extra) {
			return '<div class="fbt-card' + (extra || '') + '"><div class="fbt-card-label">' + esc(label) + '</div>' +
				'<div class="fbt-card-value">' + value + '</div><div class="fbt-card-note">' + note + '</div></div>';
		}
		var html = card('Jumlah data', fmt(rows.length), rows.length ? 'baris sesuai filter &amp; pencarian' : 'tidak ada data');
		if (cols.money > -1) {
			var sum = 0;
			rows.forEach(function (r) { var v = toNum(text(r.cells[cols.money])); if (!isNaN(v)) sum += v; });
			html += card('Total ' + cols.moneyLabel, fmt(sum), 'jumlah kolom ' + esc(cols.moneyLabel), ' accent');
		}
		if (cols.status > -1) {
			var st = {};
			rows.forEach(function (r) { var t = text(r.cells[cols.status]) || '-'; st[t] = (st[t] || 0) + 1; });
			var keys = Object.keys(st).sort(function (a, b) { return st[b] - st[a]; });
			html += card('Status', keys.length ? fmt(keys.length) + ' jenis' : '-',
				'<div class="rpt-mini">' + keys.slice(0, 4).map(function (k) { return '<span class="rpt-badge ' + stateClass(k) + '">' + esc(k) + ' ' + st[k] + '</span>'; }).join('') + '</div>');
		}
		summary.innerHTML = html;
	}

	function run() {
		if (!global.FbpTable) return;
		var list = document.querySelectorAll('table#product-table, table.fbt-auto');
		Array.prototype.forEach.call(list, function (t, idx) {
			try {
				var custom = !!t.getAttribute('data-fbt');           // halaman sudah punya tampilan sendiri
				var summary = null, cols = null;
				var onChange = function (rows) {
					if (!cols) {
						cols = pickColumns(t, rows);
						badgeStatus(rows, cols.status);
						if (idx === 0 && !custom) layoutColumns(t, rows);
					}
					renderSummary(summary, t, rows, cols);
				};
				var ok = enhance(t, onChange);
				if (idx === 0 && !custom) {
					summary = skin(t, ok);
					if (ok) {
						// render pertama terjadi sebelum kartu ringkasan dibuat
						var api = t.querySelector('tbody');
						var visible = Array.prototype.filter.call(api.rows, function (r) { return !r.classList.contains('fbt-filter') && !r.classList.contains('fbt-auto-empty'); });
						onChange(visible.filter(function (r) { return r.cells.length > 1; }));
					} else if (summary) {
						summary.parentNode.removeChild(summary);   // tabel bersubtotal: tanpa kartu ringkasan
					}
				}
			} catch (e) { /* tabel dibiarkan apa adanya */ }
		});
	}

	// ---------------------------------------------------------------------
	// Form filter laporan (REPORT: table#form-table): label + isian disusun ulang jadi kartu
	// filter dengan grid. Isian hanya dipindah (tetap di form yang sama, nama field & event
	// tidak berubah); input hidden ikut dipertahankan. Dropdown memakai pencarian dari
	// js/fbp-shell.js (body.rpt-form).
	// ---------------------------------------------------------------------
	// isian = input/select/textarea, atau wadah ber-id yang diisi dropdown lewat AJAX
	// (mis. <div id="divisilist"></div> di Data Karyawan)
	function isControl(el) { return el.querySelector('input:not([type="hidden"]), select, textarea, div[id], span[id]'); }
	function onlyButtons(cell) {
		var c = cell.querySelectorAll('input:not([type="hidden"]), select, textarea, button');
		if (!c.length) return false;
		return Array.prototype.every.call(c, function (x) {
			return x.tagName === 'BUTTON' || /^(submit|button|reset|image)$/i.test(x.type || '');
		});
	}
	function moveChildren(from, to) { while (from.firstChild) to.appendChild(from.firstChild); }

	function enhanceFilters() {
		var body = document.body;
		if (!body || /(^|\s)dash(\s|$)/.test(body.className) || body.getAttribute('data-rpt-filter') === 'off') return;
		var tables = Array.prototype.filter.call(document.querySelectorAll('table[id="form-table"]'), function (t) {
			return !t.closest('.rpt-filter') && isControl(t);
		});
		if (!tables.length) return;
		// kelompokkan per form (satu kartu per form)
		var groups = [];
		// form pemilik isian: <form> di dalam <table> (HTML tidak valid, mis. REPORT/p_Hk_Khl.asp) tidak
		// membungkus tabel, tetapi input tetap terhubung lewat input.form saat parsing
		function ownerForm(t) {
			var f = t.closest('form');
			if (f) return f;
			var ctl = t.querySelector('input, select, textarea, button');
			return ctl && ctl.form ? ctl.form : null;
		}
		tables.forEach(function (t) {
			var key = ownerForm(t) || t.parentNode, g = null;
			groups.forEach(function (x) { if (x.key === key) g = x; });
			if (!g) { g = { key: key, tables: [] }; groups.push(g); }
			g.tables.push(t);
		});
		groups.forEach(function (g) {
			try {
				var card = document.createElement('div');
				card.className = 'rpt-filter';
				var grid = document.createElement('div');
				grid.className = 'rpt-filter-grid';
				var acts = document.createElement('div');
				acts.className = 'rpt-filter-actions';
				var hidden = document.createElement('div');
				hidden.style.display = 'none';
				var nFields = 0;

				function addField(label, cell) {
					var f = document.createElement('div');
					f.className = 'rpt-field';
					if (cell.querySelector('select[multiple], textarea') || (cell.querySelector('select') && cell.querySelector('select').options.length > 25)) f.className += ' rpt-field-wide';
					var lb = document.createElement('label');
					lb.textContent = label || ' ';
					var ctl = cell.querySelector('input:not([type="hidden"]), select, textarea');
					if (ctl && ctl.id) lb.htmlFor = ctl.id;
					f.appendChild(lb);
					var box = document.createElement('div');
					box.className = 'rpt-field-ctl';
					moveChildren(cell, box);
					f.appendChild(box);
					grid.appendChild(f);
					nFields++;
				}

				g.tables.forEach(function (t) {
					var rows = Array.prototype.slice.call(t.rows);
					for (var i = 0; i < rows.length; i++) {
						var r = rows[i], cs = Array.prototype.slice.call(r.cells);
						var next = rows[i + 1];
						// pola "baris judul (th) di atas baris isian"
						var allTh = cs.length > 1 && cs.every(function (c) { return c.tagName === 'TH'; });
						if (allTh && next && isControl(next)) {
							var labels = cs.map(function (c) { return text(c); });
							Array.prototype.forEach.call(next.cells, function (c, k) {
								if (onlyButtons(c)) { moveChildren(c, acts); return; }
								if (isControl(c)) addField(labels[k] || '', c);
								else moveChildren(c, hidden);
							});
							i++;
							continue;
						}
						// pola "label lalu isian" berdampingan
						var pending = '';
						cs.forEach(function (c) {
							if (onlyButtons(c)) { moveChildren(c, acts); return; }
							if (isControl(c)) { addField(pending, c); pending = ''; return; }
							var tx = text(c);
							if (tx) pending = tx.replace(/\s*:\s*$/, '');
							moveChildren(c, hidden);   // simpan input hidden / skrip di sel ini
						});
					}
				});
				if (!nFields) return;
				card.appendChild(grid);
				if (acts.childNodes.length) card.appendChild(acts);
				card.appendChild(hidden);
				var first = g.tables[0];
				first.parentNode.insertBefore(card, first);
				// Perbaikan: bila form ada di dalam tabel, form dipindah keluar & kartu dimasukkan ke dalamnya.
				// Dulu form kosong itu ikut terhapus bersama tabel dan isian yang dipindah kehilangan formnya,
				// sehingga tombol View tidak mengirim apa pun.
				var owner = g.key && g.key.tagName === 'FORM' ? g.key : null;   // dihitung sebelum isian dipindah
				if (owner && !owner.contains(card)) {
					card.parentNode.insertBefore(owner, card);
					owner.appendChild(card);
				}
				g.tables.forEach(function (t) {
					// sisa isi tabel (mis. input hidden di luar sel) tetap dipertahankan
					Array.prototype.forEach.call(t.querySelectorAll('input, select, textarea, script'), function (x) { hidden.appendChild(x); });
					t.parentNode.removeChild(t);
				});
				// pilihan pertama terkadang kosong tanpa teks: beri teks agar jelas
				Array.prototype.forEach.call(grid.querySelectorAll('select'), function (s) {
					if (s.options.length && s.options[0].value === '' && !text(s.options[0])) s.options[0].textContent = '-- Semua --';
				});
			} catch (e) { /* form dibiarkan apa adanya */ }
		});
		if (document.querySelector('.rpt-filter')) {
			if (body.className.indexOf('rpt-form') < 0) body.className += ' rpt-form';   // dropdown dengan kolom cari
			if (body.className.indexOf('rpt-filtered') < 0) body.className += ' rpt-filtered';
		}
	}
	if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', enhanceFilters);
	else enhanceFilters();

	// ---------------------------------------------------------------------
	// Tabel daftar di halaman bertampilan baru (body.dash.page-legacy, mis. TableStatistik.asp):
	// header lama <tr bgcolor="#0C5B24"> (sel <td>) atau baris <th>, diikuti baris data.
	// Header dipindah ke <thead>, baris non-data di ujung ke <tfoot>, lalu FbpTable dipasang.
	// Dilewati: halaman dokumen/BKM/approval, tabel berisi isian (grid entri), tabel yang sudah
	// punya FbpTable sendiri, dan tabel dengan baris non-data (subtotal/judul grup) di tengah.
	// Link Edit/Hapus/detail tetap di tempatnya; cari/urut/halaman hanya menyusun tampilan.
	// ---------------------------------------------------------------------
	function isHeadRow(r) {
		if (cells(r) < 2) return false;
		if (/^#?0c5b24$/i.test(r.getAttribute('bgcolor') || '')) return true;
		return Array.prototype.every.call(r.cells, function (c) { return c.tagName === 'TH'; });
	}
	function enhanceDashList(table) {
		if (table.getAttribute('data-fbt') || table.getAttribute('data-fbt-auto') || table.closest('[data-fbt-auto="off"]')) return false;
		if (table.querySelector('table')) return false;                       // tabel bersarang (layout)
		var rows = Array.prototype.slice.call(table.rows);
		var h = -1;
		for (var i = 0; i < rows.length && i < 6; i++) { if (isHeadRow(rows[i])) { h = i; break; } }
		if (h < 0) return false;
		for (var p = 0; p < h; p++) { if ((rows[p].textContent || '').trim() || rows[p].querySelector('input:not([type="hidden"]),select,img')) return false; }
		var head = rows[h], n = cells(head), nSpan = spanOf(head);
		var body = rows.slice(h + 1).filter(function (r) { return cells(r) > 0 || (r.textContent || '').trim(); });
		var isData = body.map(function (r) { return (cells(r) === n || cells(r) === nSpan) && spanOf(r) === nSpan; });
		var first = isData.indexOf(true), last = isData.lastIndexOf(true);
		if (first !== 0 || last < 1) return false;                              // minimal 2 baris data
		for (var b = first; b <= last; b++) { if (!isData[b]) return false; }
		var dataRows = body.slice(0, last + 1);
		if (dataRows.some(function (r) { return r.querySelector('input[type="text"], input:not([type]), select, textarea'); })) return false;
		// header -> thead (baris kosong sebelum header dibuang)
		for (var e = 0; e < h; e++) rows[e].parentNode.removeChild(rows[e]);
		var thead = table.tHead || table.createTHead();
		thead.appendChild(head);
		var tbody = table.tBodies[0];
		if (!tbody) { tbody = document.createElement('tbody'); table.appendChild(tbody); }
		dataRows.forEach(function (r) { if (r.parentNode !== tbody) tbody.appendChild(r); });
		var trailing = body.slice(last + 1);
		if (trailing.length) {
			var tf = table.tFoot || table.createTFoot();
			trailing.forEach(function (r) { tf.appendChild(r); });
		}
		// tbody lain yang kini kosong dibuang (header lama kadang berada di tbody terpisah)
		Array.prototype.slice.call(table.tBodies).forEach(function (tb) { if (tb !== tbody && !tb.rows.length) tb.parentNode.removeChild(tb); });
		table.setAttribute('data-fbt-auto', 'list');
		table.className += ' fbt-dash-list';
		var resp = { close: function () {} };
		global.FbpTable.init(table, { pageSize: 20, onChange: function () { resp.close(); } });
		// toolbar (jumlah baris + cari) di atas tabel, tidak ikut melebar bersama tabel yang lebar
		var tbRow = table.querySelector('tr.fbt-toolbar-row'), tools = tbRow && tbRow.querySelector('.fbt-tools');
		if (tools) {
			var wrap = document.createElement('div');
			wrap.className = 'fbt-dash-tools';
			wrap.appendChild(tools);
			table.parentNode.insertBefore(wrap, table);
			tbRow.parentNode.removeChild(tbRow);
		}
		// wadah gulir ala DataTables scrollY: tabel panjang (mis. tampil 100 baris) digulir di dalam kotak,
		// judul kolom menempel di atas, toolbar & pagination tetap terlihat
		var box = document.createElement('div');
		box.className = 'fbt-dash-scroll';
		table.parentNode.insertBefore(box, table);
		box.appendChild(table);
		resp = responsive(table, head, nSpan);
		return true;
	}
	// Responsif ala DataTables Responsive: kolom yang tidak muat disembunyikan dari kanan
	// (NO, tombol Edit/Hapus, dan 2 kolom teks pertama selalu tampil). Klik sel pertama baris
	// membuka baris detail berisi kolom yang tersembunyi. Cetak & Export tetap semua kolom.
	var respSeq = 0;
	function responsive(table, head, nSpan) {
		var none = { close: function () {} };
		var tbody = table.tBodies[0];
		if (!tbody || !tbody.rows.length) return none;
		var byCol = Array.prototype.every.call(tbody.rows, function (r) { return r.cells.length === nSpan; });
		// kolom per sel judul: {cell, label, idx (nth-child baris data), keep}
		var cols = [], acc = 0, texts = 0;
		Array.prototype.forEach.call(head.cells, function (c, i) {
			var label = (c.textContent || '').replace(/\s+/g, ' ').trim();
			var keep = i === 0 || !label || texts < 2;
			if (label && i > 0) texts++;
			var span = c.colSpan || 1;
			for (var k = 0; k < span; k++) cols.push({ cell: c, label: label, idx: (byCol ? acc + k : i) + 1, keep: keep });
			acc += span;
		});
		if (!byCol && acc !== head.cells.length) return none;
		var hideable = cols.filter(function (c) { return !c.keep && c.label; });
		if (!hideable.length) return none;
		var id = 'fbt-r' + (++respSeq);
		table.setAttribute('data-fbt-resp', id);
		var wrap = table.parentNode;   // .fbt-dash-scroll (dibuat di enhanceDashList)
		var style = document.createElement('style');
		document.head.appendChild(style);
		var hidden = [];
		function apply() {
			var sel = hidden.map(function (c) { return 'table[data-fbt-resp="' + id + '"] > tbody > tr:not(.fbt-rdetail) > :nth-child(' + c.idx + ')'; });
			style.textContent = sel.length ? '@media screen { ' + sel.join(', ') + ' { display: none; } }' : '';
			cols.forEach(function (c) { c.cell.classList.toggle('fbt-rhide', hidden.indexOf(c) !== -1); });
			table.classList.toggle('fbt-resp-on', hidden.length > 0);
		}
		function close() {
			Array.prototype.slice.call(tbody.querySelectorAll('tr.fbt-rdetail')).forEach(function (d) { d.parentNode.removeChild(d); });
			Array.prototype.forEach.call(tbody.querySelectorAll('tr.fbt-ropen'), function (r) { r.classList.remove('fbt-ropen'); });
		}
		var lastW = -1;
		function fit() {
			var w = wrap.clientWidth;
			if (!w || w === lastW) return;
			lastW = w;
			close();
			hidden = [];
			apply();
			table.classList.remove('fbt-resp-wrap');
			// Lebar kolom diukur sekali lalu kolom yang disembunyikan dihitung langsung (dulu: sembunyikan satu
			// kolom, ukur ulang seluruh tabel, ulangi - ±1,6 detik untuk 2.000+ baris x 46 kolom).
			var over = table.offsetWidth - w;
			if (over > 1) {
				var row = null;
				for (var r = 0; r < tbody.rows.length && !row; r++) {
					var cand = tbody.rows[r];
					if (cand.offsetWidth && !cand.classList.contains('fbt-rdetail')) row = cand;
				}
				for (var i = hideable.length - 1; i >= 0 && over > 1; i--) {
					var c = hideable[i], cw = c.cell.colSpan > 1 ? 0 : c.cell.offsetWidth;
					var dc = row && row.cells[c.idx - 1];
					if (dc) cw = Math.max(cw, dc.offsetWidth);
					hidden.push(c);
					over -= cw;
				}
				apply();
				// koreksi bila perkiraan masih kurang (jarang): lanjutkan satu per satu seperti sebelumnya
				for (var k = hideable.length - 1 - hidden.length; k >= 0 && table.offsetWidth > w + 1; k--) { hidden.push(hideable[k]); apply(); }
			}
			// layar sempit: kolom yang tersisa boleh membungkus teks
			if (table.offsetWidth > w + 1) table.classList.add('fbt-resp-wrap');
		}
		tbody.addEventListener('click', function (e) {
			if (!hidden.length || e.target.closest('a, button, input, select, textarea')) return;
			var td = e.target.closest('td'), tr = td && td.parentNode;
			if (!tr || tr.parentNode !== tbody || tr.classList.contains('fbt-rdetail') || td !== tr.cells[0]) return;
			var next = tr.nextElementSibling;
			if (tr.classList.contains('fbt-ropen')) {
				tr.classList.remove('fbt-ropen');
				if (next && next.classList.contains('fbt-rdetail')) next.parentNode.removeChild(next);
				return;
			}
			var d = document.createElement('tr'), cell = document.createElement('td'), box = document.createElement('div');
			box.className = 'fbt-rd';
			d.className = 'fbt-rdetail';
			cell.colSpan = nSpan;
			hidden.slice().reverse().forEach(function (c) {
				var src = tr.cells[c.idx - 1];
				if (!src) return;
				var item = document.createElement('div'), lb = document.createElement('b'), val = document.createElement('span');
				lb.textContent = c.label;
				val.innerHTML = (src.textContent || '').trim() ? src.innerHTML : '&ndash;';
				item.appendChild(lb); item.appendChild(val); box.appendChild(item);
			});
			cell.appendChild(box); d.appendChild(cell);
			tr.parentNode.insertBefore(d, tr.nextSibling);
			tr.classList.add('fbt-ropen');
		});
		var t = null;
		function later() { clearTimeout(t); t = setTimeout(fit, 120); }
		if (global.ResizeObserver) new ResizeObserver(later).observe(wrap);
		else global.addEventListener('resize', later);
		// tabel bisa melebar setelah fit (mis. web font selesai dimuat): hitung ulang bila melewati wadah
		function recheck() { if (table.offsetWidth > wrap.clientWidth + 1) { lastW = -1; later(); } }
		if (global.ResizeObserver) new ResizeObserver(recheck).observe(table);
		if (document.fonts && document.fonts.ready) document.fonts.ready.then(function () { lastW = -1; later(); });
		fit();
		return { close: close };
	}
	function runDash() {
		var body = document.body;
		if (!global.FbpTable || !body || !/(^|\s)dash(\s|$)/.test(body.className)) return;
		if (/(^|\s)(page-doc|page-bkm|page-approval)(\s|$)/.test(body.className)) return;
		Array.prototype.forEach.call(document.querySelectorAll('.box_both table.data, .box_bothall table.data'), function (t) {
			try { enhanceDashList(t); } catch (err) { /* tabel dibiarkan apa adanya */ }
		});
	}

	// Setelah semua skrip halaman (termasuk inisialisasi FbpTable milik halaman sendiri) selesai
	if (document.readyState === 'complete') setTimeout(runDash, 0);
	else global.addEventListener('load', function () { setTimeout(runDash, 0); });
	if (document.readyState === 'complete') setTimeout(run, 0);
	else global.addEventListener('load', function () { setTimeout(run, 0); });
	global.FbpTableAuto = { run: run, enhance: enhance, runDash: runDash };
})(window);
