/* =====================================================================
   fbp-xlsx.js - export tabel HTML ke .xlsx dengan cepat, tanpa pustaka luar.
   Dibuat untuk tabel export yang besar (puluhan ribu baris): isi sel dibaca langsung
   (textContent), lembar kerja ditulis sebagai XML, lalu dizip dengan CompressionStream
   bawaan browser. Jauh lebih cepat dari tableToExcel.js (yang menghitung style per sel).

   Atribut sel yang dipahami (sama dengan tableToExcel.js):
     data-t="n"  angka   (koma pemisah ribuan dibuang)
     data-t="d"  tanggal MM/DD/YYYY -> tanggal Excel
     colspan     sel digabung (merge)
   Baris di <thead> ditulis tebal.

   Pemakaian:
     FbpXlsx.fromTable(tableEl, 'nama.xlsx', 'Sheet1').then(function () { ... });
   Bila browser tidak mendukung CompressionStream, kembalikan Promise yang ditolak
   (pemanggil bisa memakai tableToExcel.js sebagai cadangan).
   ===================================================================== */
(function (global) {
	'use strict';
	var enc = new TextEncoder();

	// ---------- CRC32 ----------
	var CRC = (function () {
		var t = new Uint32Array(256);
		for (var n = 0; n < 256; n++) {
			var c = n;
			for (var k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
			t[n] = c >>> 0;
		}
		return t;
	})();
	function crc32(bytes) {
		var c = 0xFFFFFFFF;
		for (var i = 0; i < bytes.length; i++) c = CRC[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
		return (c ^ 0xFFFFFFFF) >>> 0;
	}

	function deflate(bytes) {
		var cs = new CompressionStream('deflate-raw');
		var stream = new Blob([bytes]).stream().pipeThrough(cs);
		return new Response(stream).arrayBuffer().then(function (b) { return new Uint8Array(b); });
	}

	// ---------- ZIP (metode deflate) ----------
	function zip(files) {   // files: [{name, data(Uint8Array)}]
		return Promise.all(files.map(function (f) { return deflate(f.data); })).then(function (packed) {
			var parts = [], central = [], offset = 0;
			files.forEach(function (f, i) {
				var name = enc.encode(f.name), comp = packed[i], crc = crc32(f.data);
				var lh = new DataView(new ArrayBuffer(30));
				lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0x0800, true);
				lh.setUint16(8, 8, true); lh.setUint16(10, 0, true); lh.setUint16(12, 0x21, true);
				lh.setUint32(14, crc, true); lh.setUint32(18, comp.length, true); lh.setUint32(22, f.data.length, true);
				lh.setUint16(26, name.length, true); lh.setUint16(28, 0, true);
				parts.push(new Uint8Array(lh.buffer), name, comp);
				var ch = new DataView(new ArrayBuffer(46));
				ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true); ch.setUint16(8, 0x0800, true);
				ch.setUint16(10, 8, true); ch.setUint16(12, 0, true); ch.setUint16(14, 0x21, true);
				ch.setUint32(16, crc, true); ch.setUint32(20, comp.length, true); ch.setUint32(24, f.data.length, true);
				ch.setUint16(28, name.length, true); ch.setUint32(42, offset, true);
				central.push(new Uint8Array(ch.buffer), name);
				offset += 30 + name.length + comp.length;
			});
			var cdSize = central.reduce(function (s, p) { return s + p.length; }, 0);
			var end = new DataView(new ArrayBuffer(22));
			end.setUint32(0, 0x06054b50, true); end.setUint16(8, files.length, true); end.setUint16(10, files.length, true);
			end.setUint32(12, cdSize, true); end.setUint32(16, offset, true);
			return new Blob(parts.concat(central, [new Uint8Array(end.buffer)]), { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
		});
	}

	// ---------- XML lembar kerja ----------
	function esc(s) {
		return s.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
	}
	function colName(i) { var s = ''; i++; while (i > 0) { var m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; }
	var EPOCH = Date.UTC(1899, 11, 30);
	function excelDate(text) {
		var m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(text);
		return m ? (Date.UTC(+m[3], +m[1] - 1, +m[2]) - EPOCH) / 86400000 : null;
	}

	function sheetXml(table, sst) {   // sst: { map, list } tabel teks bersama (sharedStrings)
		var out = [], merges = [], cols = [], maxCol = 0, r = 0;
		var rows = table.rows;
		for (var ri = 0; ri < rows.length; ri++) {
			var tr = rows[ri];
			if (tr.style && tr.style.display === 'none') continue;
			var bold = tr.parentNode && tr.parentNode.tagName === 'THEAD';
			r++;
			var c = 0, cells = [];
			for (var ci = 0; ci < tr.cells.length; ci++) {
				var td = tr.cells[ci], span = td.colSpan || 1;
				while (cols[c] && cols[c] > r) c++;   // tidak dipakai (tanpa rowspan), jaga indeks
				var ref = colName(c) + r, text = (td.textContent || '').replace(/\s+/g, ' ').trim();
				var t = td.getAttribute('data-t'), cell = null;
				if (t === 'n' && text !== '') {
					var v = Number(text.replace(/,/g, ''));
					cell = isFinite(v) ? '<c r="' + ref + '"' + (bold ? ' s="2"' : '') + '><v>' + v + '</v></c>' : null;
				} else if (t === 'd' && text !== '') {
					var d = excelDate(text);
					cell = d !== null ? '<c r="' + ref + '" s="1"><v>' + d + '</v></c>' : null;
				}
				if (cell === null) {
					if (text === '') cell = '';
					else {
						var idx = sst.map.get(text);
						if (idx === undefined) { idx = sst.list.length; sst.map.set(text, idx); sst.list.push(text); }
						cell = '<c r="' + ref + '" t="s"' + (bold ? ' s="2"' : '') + '><v>' + idx + '</v></c>';
					}
				}
				cells.push(cell);
				if (span > 1) merges.push(ref + ':' + colName(c + span - 1) + r);
				c += span;
			}
			if (c > maxCol) maxCol = c;
			out.push('<row r="' + r + '">' + cells.join('') + '</row>');
		}
		return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
			'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
			(maxCol ? '<dimension ref="A1:' + colName(maxCol - 1) + Math.max(r, 1) + '"/>' : '') +
			'<sheetData>' + out.join('') + '</sheetData>' +
			(merges.length ? '<mergeCells count="' + merges.length + '">' + merges.map(function (m) { return '<mergeCell ref="' + m + '"/>'; }).join('') + '</mergeCells>' : '') +
			'</worksheet>';
	}

	function xmlEscAttr(s) { return String(s).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;'); }

	function fromTable(table, filename, sheetName) {
		if (!table || typeof CompressionStream === 'undefined' || typeof TextEncoder === 'undefined') {
			return Promise.reject(new Error('xlsx cepat tidak didukung browser ini'));
		}
		var sst = { map: new Map(), list: [] }, sheet = sheetXml(table, sst);
		var name = (sheetName || 'Sheet1').replace(/[\\\/\?\*\[\]:]/g, ' ').slice(0, 31);
		var files = [
			{ name: '[Content_Types].xml', xml: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>' },
			{ name: '_rels/.rels', xml: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>' },
			{ name: 'xl/workbook.xml', xml: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="' + xmlEscAttr(name) + '" sheetId="1" r:id="rId1"/></sheets></workbook>' },
			{ name: 'xl/_rels/workbook.xml.rels', xml: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>' },
			// gaya: 0 = biasa, 1 = tanggal (format bawaan 14), 2 = tebal
			{ name: 'xl/styles.xml', xml: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>' },
			{ name: 'xl/worksheets/sheet1.xml', xml: sheet },
			{ name: 'xl/sharedStrings.xml', xml: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' + sst.list.length + '" uniqueCount="' + sst.list.length + '">' +
				sst.list.map(function (t) { return '<si><t xml:space="preserve">' + esc(t) + '</t></si>'; }).join('') + '</sst>' }
		];
		files.forEach(function (f) { f.data = enc.encode(f.xml); f.xml = null; });
		return zip(files).then(function (blob) {
			var a = document.createElement('a');
			a.href = URL.createObjectURL(blob);
			a.download = filename || 'export.xlsx';
			document.body.appendChild(a);
			a.click();
			setTimeout(function () { URL.revokeObjectURL(a.href); a.parentNode.removeChild(a); }, 2000);
			return blob;
		});
	}

	global.FbpXlsx = { fromTable: fromTable };
})(window);
