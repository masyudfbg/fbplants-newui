/* =====================================================================
   fbp-report.js - pencarian laporan di Report/Dashboard.asp.
   Hanya menyembunyikan/menampilkan link yang sudah ada; tidak ada
   request ke server. Pasangan css/fbp-report.css.
   ===================================================================== */
(function () {
	function init() {
		var body = document.body;
		if (!body || !/(^|\s)rpt(\s|$)/.test(body.className)) return;
		// Halaman form filter (d_DashFormShow.asp) tidak butuh pencarian katalog
		if (/(^|\s)rpt-form(\s|$)/.test(body.className)) return;
		var head = document.querySelector('#siteInfo .head');
		var colL = document.querySelector('#container2 > .colL');
		if (!head || !colL) return;

		// Judul halaman di atas daftar
		var title = document.createElement('div');
		title.className = 'rpt-title';
		title.innerHTML = '<h2>Menu Laporan</h2><p>Pilih laporan di bawah, atau cari dengan kolom di atas.</p>';
		colL.insertBefore(title, colL.firstChild);

		var empty = document.createElement('div');
		empty.className = 'rpt-empty rpt-hidden';
		colL.appendChild(empty);

		// Kolom cari di header (sebelum tombol navigasi)
		var box = document.createElement('div');
		box.className = 'rpt-search';
		box.innerHTML = '<i class="fa fa-search" aria-hidden="true"></i>' +
			'<input type="search" placeholder="Cari laporan..." autocomplete="off" aria-label="Cari laporan">';
		var nav = document.getElementById('globalNav');
		head.insertBefore(box, nav);
		var input = box.querySelector('input');

		var links = Array.prototype.slice.call(colL.querySelectorAll('.navlist a, .navlist2 a'));
		links.forEach(function (a) {
			a.setAttribute('data-label', a.textContent);
			a.innerHTML = '<span>' + esc(a.textContent) + '</span>';
		});
		var cards = Array.prototype.slice.call(colL.querySelectorAll('.navlist, .navlist2'));
		var headers = Array.prototype.slice.call(colL.querySelectorAll('.entry-header'));
		var bodies = Array.prototype.slice.call(colL.querySelectorAll('.entry-body'));
		var jump = document.querySelectorAll('#logo-scroll li');

		function esc(s) { return s.replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }

		function mark(text, tokens) {
			if (!tokens.length) return esc(text);
			var lower = text.toLowerCase(), hit = [];
			tokens.forEach(function (t) {
				var from = 0, at;
				while ((at = lower.indexOf(t, from)) !== -1) { for (var k = at; k < at + t.length; k++) hit[k] = 1; from = at + t.length; }
			});
			var out = '', open = false;
			for (var i = 0; i < text.length; i++) {
				if (hit[i] && !open) { out += '<mark>'; open = true; }
				if (!hit[i] && open) { out += '</mark>'; open = false; }
				out += esc(text.charAt(i));
			}
			return out + (open ? '</mark>' : '');
		}

		function hasVisibleLink(el) {
			var as = el.querySelectorAll('a');
			for (var i = 0; i < as.length; i++) if (!as[i].parentNode.classList.contains('rpt-hidden')) return true;
			return false;
		}

		function apply() {
			var q = input.value.trim().toLowerCase();
			var tokens = q ? q.split(/\s+/) : [];
			var total = 0;
			links.forEach(function (a) {
				var label = a.getAttribute('data-label');
				var lower = label.toLowerCase();
				var ok = tokens.every(function (t) { return lower.indexOf(t) !== -1; });
				a.parentNode.classList.toggle('rpt-hidden', !ok);
				a.innerHTML = '<span>' + mark(label, ok ? tokens : []) + '</span>';
				if (ok) total++;
			});
			// Kartu tanpa link yang cocok disembunyikan (hanya saat mencari)
			cards.forEach(function (c) { c.classList.toggle('rpt-hidden', !!tokens.length && !hasVisibleLink(c)); });
			// Kategori: entry-header diikuti satu atau lebih entry-body
			var visibleIds = {};
			headers.forEach(function (h) {
				var any = false, el = h.nextElementSibling;
				while (el && !el.classList.contains('entry-header')) {
					if (el.classList.contains('entry-body')) {
						var show = !tokens.length || hasVisibleLink(el);
						el.classList.toggle('rpt-hidden', !show);
						any = any || show;
					}
					el = el.nextElementSibling;
				}
				h.classList.toggle('rpt-hidden', !any);
				if (any) visibleIds[h.id] = true;
			});
			bodies.forEach(function (b) {
				// entry-body yang tidak berada tepat setelah header (struktur bersarang)
				if (!tokens.length) b.classList.remove('rpt-hidden');
			});
			Array.prototype.forEach.call(jump, function (li) {
				var a = li.querySelector('a');
				var id = a ? (a.getAttribute('href') || '').replace('#', '') : '';
				li.classList.toggle('rpt-hidden', !!tokens.length && id && !visibleIds[id]);
			});
			empty.classList.toggle('rpt-hidden', !tokens.length || total > 0);
			if (tokens.length && !total) empty.textContent = 'Tidak ada laporan yang cocok dengan "' + input.value.trim() + '".';
		}

		input.addEventListener('input', apply);
		input.addEventListener('keydown', function (e) {
			if (e.key === 'Escape') { input.value = ''; apply(); }
			if (e.key === 'Enter') {
				e.preventDefault();
				// Enter membuka laporan pertama yang cocok
				for (var i = 0; i < links.length; i++) {
					if (!links[i].parentNode.classList.contains('rpt-hidden')) { links[i].click(); break; }
				}
			}
		});
		document.addEventListener('keydown', function (e) {
			var tag = (e.target.tagName || '').toLowerCase();
			if (((e.ctrlKey || e.metaKey) && (e.key === 'k' || e.key === 'K')) || (e.key === '/' && tag !== 'input' && tag !== 'textarea')) {
				e.preventDefault(); input.focus(); input.select();
			}
		});
	}

	if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
	else init();
})();
