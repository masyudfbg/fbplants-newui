/* =====================================================================
   Cetak Bukti Jurnal (bkm_jurnal_view.asp): header + detail jurnal.
   Data diambil dari halaman yang sedang tampil (tidak ada query baru),
   lalu dicetak lewat iframe tersembunyi (tidak terhalang pemblokir popup).
   ===================================================================== */
(function (global) {
	'use strict';
	function txt(el) { return el ? (el.value !== undefined && el.tagName !== 'TD' ? el.value : el.textContent).replace(/\s+/g, ' ').trim() : ''; }
	function esc(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
	function byId(id) { return document.getElementById(id); }

	// nilai di sel sebelah label (mis. "Posting" -> "YA")
	function valueAfterLabel(label) {
		var cells = document.querySelectorAll('td');
		for (var i = 0; i < cells.length; i++) {
			if (txt(cells[i]).toLowerCase() === label.toLowerCase() && cells[i].nextElementSibling) return txt(cells[i].nextElementSibling);
		}
		return '';
	}

	function collect() {
		var title = txt(document.querySelector('#head_judul strong')).replace(/\s*-\s*MODULE\s*:.*$/i, '');
		var parts = title.split(/\s+-\s+/);
		var d = {
			company: parts[0] || '',
			ba: parts[1] || '',
			no: txt(byId('jurnalno')),
			source: txt(byId('sourcecode')),
			date: txt(byId('jurnaldate')),
			remark: byId('remarks') ? byId('remarks').value.trim() : '',
			posting: valueAfterLabel('Posting'),
			head: [], rows: [], totalDebit: '', totalKredit: ''
		};
		var table = byId('table2');
		if (!table) return d;
		var hr = table.querySelector('tr#trheader2') || table.rows[0];
		Array.prototype.forEach.call(hr.cells, function (c) { d.head.push(txt(c)); });
		Array.prototype.forEach.call(table.rows, function (r) {
			if (/^tru\d+$/.test(r.id || '')) {
				var cells = [];
				for (var i = 0; i < d.head.length; i++) cells.push(r.cells[i] ? txt(r.cells[i]) : '');
				d.rows.push(cells);
			} else if (r.className.indexOf('ratakanan') > -1) {
				var inp = r.querySelectorAll('input');
				if (inp.length >= 2) { d.totalDebit = inp[0].value; d.totalKredit = inp[1].value; }
			}
		});
		return d;
	}

	function buildHtml(d) {
		var numCol = function (h) { return /^(DEBIT|KREDIT|CREDIT|AMOUNT)$/i.test(h); };
		var now = new Date();
		var printed = ('0' + now.getDate()).slice(-2) + '/' + ('0' + (now.getMonth() + 1)).slice(-2) + '/' + now.getFullYear() + ' ' + ('0' + now.getHours()).slice(-2) + ':' + ('0' + now.getMinutes()).slice(-2);
		var thead = '<tr>' + d.head.map(function (h) { return '<th' + (numCol(h) ? ' class="num"' : '') + '>' + esc(h) + '</th>'; }).join('') + '</tr>';
		var body = d.rows.map(function (r) {
			return '<tr>' + r.map(function (v, i) { return '<td' + (numCol(d.head[i]) ? ' class="num"' : '') + '>' + esc(v) + '</td>'; }).join('') + '</tr>';
		}).join('') || '<tr><td colspan="' + d.head.length + '" class="empty">Tidak ada detail jurnal</td></tr>';
		var iDeb = -1, iKre = -1;
		d.head.forEach(function (h, i) { if (/^DEBIT$/i.test(h)) iDeb = i; if (/^(KREDIT|CREDIT)$/i.test(h)) iKre = i; });
		var foot = '';
		if (iDeb > -1) {
			foot = '<tr class="total"><td colspan="' + iDeb + '" class="lbl">Total</td>';
			for (var i = iDeb; i < d.head.length; i++) foot += '<td class="num">' + esc(i === iDeb ? d.totalDebit : i === iKre ? d.totalKredit : '') + '</td>';
			foot += '</tr>';
		}
		var css = [
			'@page { size: A4 landscape; margin: 12mm; }',
			'* { box-sizing: border-box; }',
			'body { margin: 0; font-family: Arial, Helvetica, sans-serif; font-size: 11px; color: #111; }',
			'.top { display: flex; justify-content: space-between; align-items: flex-end; border-bottom: 2px solid #111; padding-bottom: 6px; margin-bottom: 10px; }',
			'.co { font-size: 14px; font-weight: bold; } .ba { font-size: 11px; color: #333; margin-top: 2px; }',
			'h1 { margin: 0; font-size: 16px; letter-spacing: 1px; text-align: right; }',
			'.meta { display: grid; grid-template-columns: 110px 1fr 110px 1fr; gap: 4px 10px; margin-bottom: 10px; }',
			'.meta b { font-weight: bold; } .meta .wide { grid-column: 2 / 5; white-space: pre-wrap; }',
			'table { width: 100%; border-collapse: collapse; }',
			'th, td { border: 1px solid #555; padding: 4px 6px; text-align: left; vertical-align: top; }',
			'th { background: #eee; font-size: 10px; text-transform: uppercase; -webkit-print-color-adjust: exact; print-color-adjust: exact; }',
			'.num { text-align: right; white-space: nowrap; }',
			'tr.total td { font-weight: bold; background: #f5f5f5; -webkit-print-color-adjust: exact; print-color-adjust: exact; }',
			'tr.total .lbl { text-align: right; } .empty { text-align: center; color: #666; }',
			'thead { display: table-header-group; } tr { page-break-inside: avoid; }',
			'.sign { display: grid; grid-template-columns: repeat(3, 1fr); gap: 30px; margin-top: 28px; text-align: center; }',
			'.sign .line { margin-top: 55px; border-top: 1px solid #111; padding-top: 3px; }',
			'.foot { margin-top: 14px; font-size: 9px; color: #555; }'
		].join('\n');
		return '<!doctype html><html><head><meta charset="utf-8"><title>Bukti Jurnal ' + esc(d.no) + '</title><style>' + css + '</style></head><body>' +
			'<div class="top"><div><div class="co">' + esc(d.company) + '</div><div class="ba">' + esc(d.ba) + '</div></div><h1>BUKTI JURNAL</h1></div>' +
			'<div class="meta">' +
			'<b>No. Jurnal</b><span>' + esc(d.no) + '</span><b>Tanggal</b><span>' + esc(d.date) + '</span>' +
			'<b>Source Code</b><span>' + esc(d.source) + '</span><b>Posting</b><span>' + esc(d.posting) + '</span>' +
			'<b>Remark</b><span class="wide">' + esc(d.remark) + '</span>' +
			'</div>' +
			'<table><thead>' + thead + '</thead><tbody>' + body + '</tbody>' + (foot ? '<tfoot>' + foot + '</tfoot>' : '') + '</table>' +
			'<div class="sign"><div>Dibuat oleh<div class="line">&nbsp;</div></div><div>Diperiksa oleh<div class="line">&nbsp;</div></div><div>Disetujui oleh<div class="line">&nbsp;</div></div></div>' +
			'<div class="foot">Dicetak dari FB-Plants pada ' + esc(printed) + '</div>' +
			'</body></html>';
	}

	global.fbpPrintJurnal = function () {
		var html = buildHtml(collect());
		var old = byId('fbp-print-frame');
		if (old) old.parentNode.removeChild(old);
		var f = document.createElement('iframe');
		f.id = 'fbp-print-frame';
		f.setAttribute('aria-hidden', 'true');
		f.style.cssText = 'position:fixed;right:0;bottom:0;width:0;height:0;border:0;visibility:hidden;';
		document.body.appendChild(f);
		var doc = f.contentWindow.document;
		doc.open();
		doc.write(html);
		doc.close();
		setTimeout(function () { f.contentWindow.focus(); f.contentWindow.print(); }, 250);
	};
	// dipakai untuk pengujian / pratinjau
	global.fbpPrintJurnalHtml = function () { return buildHtml(collect()); };
})(window);
