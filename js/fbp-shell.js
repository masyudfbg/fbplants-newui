/* =====================================================================
   fbp-shell.js - tombol hide/show sidebar menu untuk halaman yang
   memakai tampilan baru (<body class="dash">). Pasangan css/fbp-shell.css.
   Dimuat di <head>; pilihan user disimpan di browser (localStorage).
   ===================================================================== */
(function () {
	var KEY = 'fbp_sidebar_collapsed';
	var root = document.documentElement;

	function load() { try { return localStorage.getItem(KEY); } catch (e) { return null; } }
	function save(v) { try { localStorage.setItem(KEY, v); } catch (e) {} }
	function isCollapsed() { return /(^|\s)fbp-sb-collapsed(\s|$)/.test(root.className); }
	function setCollapsed(on) {
		root.className = root.className.replace(/\s*fbp-sb-collapsed/g, '') + (on ? ' fbp-sb-collapsed' : '');
	}

	// Tanpa meta viewport, browser HP merender halaman selebar ~980px lalu memperkecilnya
	// sehingga tampilan HP (@media max-width 760px) tidak pernah aktif. Tambahkan bila belum ada.
	(function () {
		var head = document.head || document.getElementsByTagName('head')[0];
		if (head && !document.querySelector('meta[name="viewport"]')) {
			var m = document.createElement('meta');
			m.name = 'viewport';
			m.content = 'width=device-width, initial-scale=1';
			head.appendChild(m);
		}
	})();

	// Mode HP: menu jadi laci (overlay) yang selalu mulai tertutup; buka/tutup di HP
	// tidak disimpan agar pilihan di desktop tidak berubah.
	var mq = window.matchMedia ? window.matchMedia('(max-width: 760px)') : null;
	function isMobile() { return mq ? mq.matches : window.innerWidth <= 760; }

	// Pasang status awal sebelum <body> dirender agar sidebar tidak "berkedip".
	// Default: tampil di layar lebar, tersembunyi di layar sempit.
	var saved = load();
	setCollapsed(isMobile() ? true : (saved === null ? window.innerWidth < 900 : saved === '1'));

	function init() {
		var body = document.body;
		var header = document.getElementById('siteInfo');
		if (!body || !header || !/(^|\s)dash(\s|$)/.test(body.className)) return;
		if (document.getElementById('fbpSidebarToggle') || header.getAttribute('data-fbp-init')) return;
		header.setAttribute('data-fbp-init', '1');

		// Label tombol navigasi kanan sebagai tooltip (saat hanya ikon yang tampil)
		var navLinks = header.querySelectorAll('#globalNav a');
		for (var n = 0; n < navLinks.length; n++) {
			var span = navLinks[n].querySelector('span');
			if (span && !navLinks[n].getAttribute('title')) navLinks[n].setAttribute('title', span.textContent);
		}

		// Halaman tanpa sidebar menu (mis. input BKM): tanpa tombol hide/show & tanpa cari menu
		if (!document.querySelector('.dtree')) {
			body.className += ' fbp-no-menu';
			return;
		}

		var btn = document.createElement('button');
		btn.type = 'button';
		btn.id = 'fbpSidebarToggle';
		btn.className = 'fbp-sidebar-toggle';
		btn.innerHTML = '<i class="fa fa-bars" aria-hidden="true"></i>';
		header.insertBefore(btn, header.firstChild);

		function sync() {
			var label = isCollapsed() ? 'Tampilkan menu' : 'Sembunyikan menu';
			btn.setAttribute('aria-label', label);
			btn.setAttribute('title', label);
			btn.setAttribute('aria-expanded', isCollapsed() ? 'false' : 'true');
		}
		sync();

		// Animasi baru aktif setelah halaman tampil
		setTimeout(function () { root.className += ' fbp-sb-anim'; }, 50);

		btn.addEventListener('click', function () {
			var on = !isCollapsed();
			setCollapsed(on);
			if (!isMobile()) save(on ? '1' : '0');
			sync();
			// Grafik (Chart.js) menyesuaikan lebar setelah animasi selesai
			setTimeout(function () {
				try { window.dispatchEvent(new Event('resize')); } catch (e) {}
			}, 260);
		});

		// Latar gelap di belakang laci menu (HP): ketuk untuk menutup menu
		var backdrop = document.createElement('div');
		backdrop.className = 'fbp-sb-backdrop';
		backdrop.setAttribute('aria-hidden', 'true');
		body.appendChild(backdrop);
		function closeDrawer() {
			if (isMobile() && !isCollapsed()) { setCollapsed(true); sync(); }
		}
		backdrop.addEventListener('click', closeDrawer);
		document.addEventListener('keydown', function (e) { if (e.key === 'Escape') closeDrawer(); });
		// Pindah dari layar lebar ke HP (putar layar / ubah ukuran): menu ditutup
		if (mq) {
			var onMq = function (e) { if (e.matches) { setCollapsed(true); sync(); } };
			if (mq.addEventListener) mq.addEventListener('change', onMq); else if (mq.addListener) mq.addListener(onMq);
		}

		initMenuSearch(header);
	}

	// =====================================================================
	// Pencarian menu di header. Sumber data: objek dTree "d" dari menu.asp
	// (menu yang memang boleh diakses user). Tidak ada request ke server.
	// =====================================================================
	function initMenuSearch(header) {
		var wrap = document.createElement('div');
		wrap.className = 'fbp-msearch';
		wrap.innerHTML =
			'<i class="fa fa-search" aria-hidden="true"></i>' +
			'<input type="text" id="fbpMenuSearch" placeholder="Cari menu..." autocomplete="off" spellcheck="false"' +
			' role="combobox" aria-expanded="false" aria-controls="fbpMenuResults" aria-autocomplete="list" aria-label="Cari menu">' +
			'<kbd aria-hidden="true">Ctrl K</kbd>' +
			'<div class="fbp-msearch-panel" id="fbpMenuResults" role="listbox"></div>';
		header.appendChild(wrap);

		var input = wrap.querySelector('input');
		var panel = wrap.querySelector('.fbp-msearch-panel');
		var index = null, items = [], active = -1;

		function esc(s) {
			return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; });
		}

		// Kumpulkan semua menu yang punya link, beserta jalur foldernya
		function buildIndex() {
			var tree = window.d;
			if (!tree || !tree.aNodes) return [];
			var byId = {}, list = [], i;
			for (i = 0; i < tree.aNodes.length; i++) byId[tree.aNodes[i].id] = tree.aNodes[i];
			for (i = 0; i < tree.aNodes.length; i++) {
				var node = tree.aNodes[i];
				if (!node.url) continue;
				// Hanya menu yang benar-benar tampil di sidebar: seluruh induknya harus ada di data menu.
				// menu.asp mengecek ISACTIVE per baris, jadi menu aktif di bawah induk yang nonaktif
				// tetap terkirim ke dTree (tidak tampil di sidebar) - jangan ikut muncul di pencarian.
				var path = [], cur = node, ok = true, guard = 0;
				while (cur.pid != -1 && guard++ < 20) {
					var p = byId[cur.pid];
					if (!p) { ok = false; break; }
					if (p.pid != -1) path.unshift(p.name);
					cur = p;
				}
				if (!ok) continue;
				var name = String(node.name).replace(/<[^>]*>/g, '');
				list.push({ name: name, path: path.join(' / '), url: node.url, hay: (name + ' ' + path.join(' ')).toLowerCase() });
			}
			return list;
		}

		function highlight(text, tokens) {
			var lower = text.toLowerCase(), marked = [];
			tokens.forEach(function (t) {
				var from = 0, at;
				while (t && (at = lower.indexOf(t, from)) !== -1) {
					for (var k = at; k < at + t.length; k++) marked[k] = true;
					from = at + t.length;
				}
			});
			var out = '', open = false;
			for (var i = 0; i < text.length; i++) {
				if (marked[i] && !open) { out += '<mark>'; open = true; }
				if (!marked[i] && open) { out += '</mark>'; open = false; }
				out += esc(text.charAt(i));
			}
			return out + (open ? '</mark>' : '');
		}

		function setActive(i) {
			if (active > -1 && items[active]) items[active].classList.remove('active');
			active = items.length ? Math.max(0, Math.min(i, items.length - 1)) : -1;
			if (active > -1) {
				var el = items[active];
				el.classList.add('active');
				input.setAttribute('aria-activedescendant', el.id);
				if (el.offsetTop < panel.scrollTop) panel.scrollTop = el.offsetTop - 6;
				else if (el.offsetTop + el.offsetHeight > panel.scrollTop + panel.clientHeight) panel.scrollTop = el.offsetTop + el.offsetHeight - panel.clientHeight + 6;
			} else input.removeAttribute('aria-activedescendant');
		}

		function open() { wrap.classList.add('open'); input.setAttribute('aria-expanded', 'true'); }
		function close() { wrap.classList.remove('open'); input.setAttribute('aria-expanded', 'false'); }

		function render() {
			if (!index) index = buildIndex();
			var q = input.value.trim().toLowerCase();
			items = []; active = -1;
			if (!q) { panel.innerHTML = ''; close(); return; }
			var tokens = q.split(/\s+/);
			var hits = index.filter(function (m) {
				return tokens.every(function (t) { return m.hay.indexOf(t) !== -1; });
			});
			// Nama menu yang diawali kata kunci tampil lebih dulu
			hits.sort(function (a, b) {
				var sa = a.name.toLowerCase().indexOf(tokens[0]) === 0 ? 0 : 1;
				var sb = b.name.toLowerCase().indexOf(tokens[0]) === 0 ? 0 : 1;
				return sa - sb || a.name.localeCompare(b.name);
			});
			var html = '';
			if (!index.length) html = '<div class="fbp-msearch-empty">Menu belum tersedia di halaman ini.</div>';
			else if (!hits.length) html = '<div class="fbp-msearch-empty">Tidak ada menu yang cocok dengan "' + esc(input.value.trim()) + '"</div>';
			else {
				html = '<div class="fbp-msearch-count">' + hits.length + ' menu ditemukan</div>';
				hits.slice(0, 50).forEach(function (m, i) {
					html += '<a class="fbp-msearch-item" role="option" id="fbpMenuOpt' + i + '" href="' + esc(m.url) + '">' +
						'<i class="fa fa-file-text-o" aria-hidden="true"></i>' +
						'<span><b>' + highlight(m.name, tokens) + '</b>' + (m.path ? '<small>' + highlight(m.path, tokens) + '</small>' : '') + '</span></a>';
				});
			}
			panel.innerHTML = html;
			items = Array.prototype.slice.call(panel.querySelectorAll('.fbp-msearch-item'));
			setActive(0);
			open();
		}

		input.addEventListener('input', render);
		input.addEventListener('focus', function () { if (input.value.trim()) render(); });
		input.addEventListener('keydown', function (e) {
			if (e.key === 'ArrowDown') { e.preventDefault(); setActive(active + 1); }
			else if (e.key === 'ArrowUp') { e.preventDefault(); setActive(active - 1); }
			else if (e.key === 'Enter') { e.preventDefault(); if (active > -1) window.location.href = items[active].getAttribute('href'); }
			else if (e.key === 'Escape') { if (input.value) { input.value = ''; render(); } else input.blur(); }
		});
		panel.addEventListener('mousemove', function (e) {
			var el = e.target.closest ? e.target.closest('.fbp-msearch-item') : null;
			if (el && items.indexOf(el) !== active) setActive(items.indexOf(el));
		});
		document.addEventListener('mousedown', function (e) { if (!wrap.contains(e.target)) close(); });

		// Pintasan: Ctrl+K atau "/" (bila tidak sedang mengetik di field lain)
		document.addEventListener('keydown', function (e) {
			var tag = (e.target.tagName || '').toLowerCase();
			var typing = tag === 'input' || tag === 'textarea' || tag === 'select' || e.target.isContentEditable;
			if (((e.ctrlKey || e.metaKey) && (e.key === 'k' || e.key === 'K')) || (e.key === '/' && !typing)) {
				e.preventDefault();
				input.focus();
				input.select();
			}
		});
	}

	// =====================================================================
	// Dropdown dengan kolom pencarian untuk semua <select> biasa di halaman
	// bertampilan baru (body.dash) dan form filter laporan (body.rpt-form). <select> asli tetap dipakai (nilai, name, onchange,
	// AJAX); hanya daftar pilihannya yang diganti panel dengan kolom cari.
	// Dilewati: multiple / size, disabled, sudah chosen / select2, opsi < 7,
	// perangkat sentuh (picker bawaan lebih nyaman).
	// =====================================================================
	var MIN_OPTIONS = 7;
	var coarse = window.matchMedia && window.matchMedia('(pointer: coarse)').matches;
	var panel = null, current = null, items = [], active = -1;

	function eligible(sel) {
		if (!sel || sel.tagName !== 'SELECT' || coarse) return false;
		if (!/(^|\s)(dash|rpt-form)(\s|$)/.test(document.body.className)) return false;
		if (sel.multiple || sel.size > 1 || sel.disabled) return false;
		if (sel.getAttribute('data-fbp-native') !== null) return false;
		if (/(^|\s)(select2-hidden-accessible|chzn-done)(\s|$)/.test(sel.className)) return false;
		var next = sel.nextElementSibling;
		if (next && /(^|\s)(chosen-container|select2|select2-container)(\s|$)/.test(next.className)) return false;
		return sel.options.length >= MIN_OPTIONS;
	}

	function esc(s) {
		return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; });
	}

	function buildPanel() {
		panel = document.createElement('div');
		panel.className = 'fbp-sel-panel';
		panel.innerHTML = '<div class="fbp-sel-search"><i class="fa fa-search" aria-hidden="true"></i>' +
			'<input type="text" autocomplete="off" spellcheck="false" placeholder="Cari..." aria-label="Cari pilihan"></div>' +
			'<div class="fbp-sel-list" role="listbox"></div>';
		document.body.appendChild(panel);
		var input = panel.querySelector('input');
		input.addEventListener('input', function () { renderList(input.value); });
		input.addEventListener('keydown', function (e) {
			if (e.key === 'ArrowDown') { e.preventDefault(); setActive(active + 1); }
			else if (e.key === 'ArrowUp') { e.preventDefault(); setActive(active - 1); }
			else if (e.key === 'Enter') { e.preventDefault(); if (active > -1) choose(items[active]); }
			else if (e.key === 'Escape' || e.key === 'Tab') { var s = current; closePanel(); if (s && e.key === 'Escape') s.focus(); }
		});
		panel.addEventListener('mousedown', function (e) {
			var it = e.target.closest ? e.target.closest('.fbp-sel-item') : null;
			if (it) { e.preventDefault(); choose(it); }
		});
		panel.addEventListener('mousemove', function (e) {
			var it = e.target.closest ? e.target.closest('.fbp-sel-item') : null;
			if (it && items.indexOf(it) !== active) setActive(items.indexOf(it));
		});
	}

	function renderList(q) {
		var list = panel.querySelector('.fbp-sel-list');
		var tokens = String(q || '').toLowerCase().split(/\s+/).filter(Boolean);
		var html = '', n = 0, lastGroup = null;
		for (var i = 0; i < current.options.length; i++) {
			var o = current.options[i];
			if (o.hidden || o.style.display === 'none') continue;
			var text = (o.text || '').replace(/\s+/g, ' ').trim();
			var hay = (text + ' ' + o.value).toLowerCase();
			if (!tokens.every(function (t) { return hay.indexOf(t) !== -1; })) continue;
			var g = o.parentNode && o.parentNode.tagName === 'OPTGROUP' ? o.parentNode.label : null;
			if (g && g !== lastGroup) html += '<div class="fbp-sel-group">' + esc(g) + '</div>';
			lastGroup = g;
			html += '<div class="fbp-sel-item' + (o.disabled ? ' disabled' : '') + (i === current.selectedIndex ? ' selected' : '') +
				'" role="option" data-i="' + i + '">' + (text ? esc(text) : '<span class="fbp-sel-empty">(kosong)</span>') + '</div>';
			n++;
		}
		if (!n) html = '<div class="fbp-sel-none">Tidak ada pilihan yang cocok</div>';
		list.innerHTML = html;
		items = Array.prototype.slice.call(list.querySelectorAll('.fbp-sel-item:not(.disabled)'));
		var sel = list.querySelector('.fbp-sel-item.selected:not(.disabled)');
		setActive(sel && !tokens.length ? items.indexOf(sel) : 0, true);
	}

	function setActive(i, center) {
		if (active > -1 && items[active]) items[active].classList.remove('active');
		active = items.length ? Math.max(0, Math.min(i, items.length - 1)) : -1;
		if (active < 0) return;
		var el = items[active], list = el.parentNode;
		el.classList.add('active');
		if (center) list.scrollTop = el.offsetTop - list.clientHeight / 2 + el.offsetHeight / 2;
		else if (el.offsetTop < list.scrollTop) list.scrollTop = el.offsetTop;
		else if (el.offsetTop + el.offsetHeight > list.scrollTop + list.clientHeight) list.scrollTop = el.offsetTop + el.offsetHeight - list.clientHeight;
	}

	function place() {
		if (!current || !panel) return;
		var r = current.getBoundingClientRect();
		if (!r.width && !r.height) { closePanel(); return; }
		var w = Math.max(r.width, 260), left = Math.min(r.left, window.innerWidth - w - 8);
		var below = window.innerHeight - r.bottom, h = Math.min(360, Math.max(below, r.top) - 12);
		panel.style.width = w + 'px';
		panel.style.left = Math.max(8, left) + 'px';
		panel.querySelector('.fbp-sel-list').style.maxHeight = Math.max(120, h - 52) + 'px';
		if (below >= 220 || below >= r.top) { panel.style.top = (r.bottom + 4) + 'px'; panel.style.bottom = 'auto'; }
		else { panel.style.top = 'auto'; panel.style.bottom = (window.innerHeight - r.top + 4) + 'px'; }
	}

	function openPanel(sel) {
		if (!panel) buildPanel();
		current = sel;
		var input = panel.querySelector('input');
		input.value = '';
		renderList('');
		panel.classList.add('open');
		place();
		input.focus();
	}

	function closePanel() {
		if (panel) panel.classList.remove('open');
		current = null; items = []; active = -1;
	}

	function choose(it) {
		if (!current || !it) return;
		var sel = current, i = parseInt(it.getAttribute('data-i'), 10);
		closePanel();
		sel.focus();
		if (sel.selectedIndex === i) return;
		sel.selectedIndex = i;
		var ev;
		try { ev = new Event('change', { bubbles: true }); } catch (x) { ev = document.createEvent('HTMLEvents'); ev.initEvent('change', true, false); }
		sel.dispatchEvent(ev);
	}

	document.addEventListener('mousedown', function (e) {
		var sel = e.target;
		if (sel && sel.tagName === 'OPTION') sel = sel.parentNode.tagName === 'SELECT' ? sel.parentNode : sel.parentNode.parentNode;
		if (eligible(sel)) {
			e.preventDefault();
			if (current === sel) { closePanel(); return; }
			sel.focus();
			openPanel(sel);
			return;
		}
		if (panel && current && !panel.contains(e.target)) closePanel();
	}, true);
	document.addEventListener('keydown', function (e) {
		var sel = e.target;
		if (!eligible(sel)) return;
		if (e.key === 'Enter' || e.key === ' ' || e.key === 'F4' || (e.altKey && (e.key === 'ArrowDown' || e.key === 'ArrowUp'))) {
			e.preventDefault();
			openPanel(sel);
		}
	}, true);
	window.addEventListener('resize', function () { if (current) place(); });
	document.addEventListener('scroll', function (e) {
		if (current && panel && !panel.contains(e.target)) place();
	}, true);

	// =====================================================================
	// Halaman dokumen (body.page-doc): link popup pilih data di sel yang sama dengan isian
	// (mis. "Stock", "Job", "Location Code") ditandai .fbp-pick agar tampil sebagai tombol
	// di samping isian. Span nama hasil pilihan (.satuan_item, diisi AJAX lewat id-nya)
	// dipindah ke bawah isian sebagai keterangan, sehingga panjang tombol tetap.
	// =====================================================================
	function initPickers() {
		if (!/(^|\s)page-doc(\s|$)/.test(document.body.className)) return;
		var links = document.querySelectorAll('td a[href^="javascript:launchwin"], td a[href^="javascript:Popup"]');
		Array.prototype.forEach.call(links, function (a) {
			var td = a.closest('td');
			if (!td || !td.querySelector('input[type="text"], input:not([type])')) return;
			if (a.querySelector('img')) return;                 // ikon cetak dsb. dibiarkan
			a.className += ' fbp-pick';
			td.className += ' fbp-pick-cell';
			Array.prototype.forEach.call(a.querySelectorAll('.satuan_item'), function (sp) {
				sp.className += ' fbp-pick-hint';
				td.appendChild(sp);
			});
		});
	}
	if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', initPickers);
	else initPickers();

	// =====================================================================
	// Tabel daftar polos (header lama <tr bgcolor="#0C5B24"> atau <th>) di halaman bertampilan
	// baru: muat js/fbp-table.js + js/fbp-table-auto.js hanya bila ada kandidat, agar tampil
	// ala DataTables (cari, jumlah baris, urut, pagination). Lihat enhanceDashList di fbp-table-auto.js.
	// =====================================================================
	function loadListTables() {
		var b = document.body;
		if (!/(^|\s)dash(\s|$)/.test(b.className) || /(^|\s)(page-doc|page-bkm|page-approval)(\s|$)/.test(b.className)) return;
		var cand = document.querySelector('.box_both table.data tr[bgcolor="#0C5B24" i], .box_both table.data tr > th, .box_bothall table.data tr[bgcolor="#0C5B24" i]');
		if (!cand) return;
		if (window.FbpTableAuto) { if (window.FbpTableAuto.runDash) window.FbpTableAuto.runDash(); return; }
		if (window.__fbtLoading) return;
		window.__fbtLoading = true;
		var css = document.createElement('link');
		css.rel = 'stylesheet';
		css.href = '/css/fbp-table.css?v=20261010c';
		document.head.appendChild(css);
		['/js/fbp-table.js?v=20261015a', '/js/fbp-table-auto.js?v=20261015a'].forEach(function (src) {
			var s = document.createElement('script');
			s.src = src;
			s.async = false;   // urutan: fbp-table.js dulu
			document.head.appendChild(s);
		});
	}
	if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', loadListTables);
	else loadListTables();

	if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
	else init();
})();
