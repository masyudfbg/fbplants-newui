/* =====================================================================
   Dropdown ganti Business Unit (BA) tanpa logout.
   Tombol: <button class="ba-switch" data-tok="..."> (login.asp, kartu "Selamat datang").
   Data & proses: ba_switch.asp (list / switch). Setelah berhasil halaman dimuat ulang.
   ===================================================================== */
(function () {
	'use strict';
	var btn = document.querySelector('.ba-switch');
	if (!btn) return;
	var panel = null, input = null, list = null, items = [], loaded = false, busy = false, active = -1;

	function el(tag, cls, text) {
		var e = document.createElement(tag);
		if (cls) e.className = cls;
		if (text != null) e.textContent = text;
		return e;
	}
	function build() {
		panel = el('div', 'ba-panel');
		panel.setAttribute('role', 'dialog');
		panel.setAttribute('aria-label', 'Ganti Business Unit');
		var head = el('div', 'ba-panel-head');
		head.appendChild(el('b', '', 'Ganti Business Unit'));
		head.appendChild(el('span', '', 'BA yang dipilih juga menjadi BA default saat login.'));
		panel.appendChild(head);
		input = el('input', 'ba-panel-search');
		input.type = 'search';
		input.placeholder = 'Cari kode / nama BA / PT...';
		input.setAttribute('aria-label', 'Cari BA');
		panel.appendChild(input);
		list = el('div', 'ba-panel-list');
		list.setAttribute('role', 'listbox');
		list.appendChild(el('div', 'ba-panel-msg', 'Memuat daftar BA...'));
		panel.appendChild(list);
		document.body.appendChild(panel);
		input.addEventListener('input', filter);
		input.addEventListener('keydown', keys);
		list.addEventListener('click', function (e) {
			var it = e.target.closest('.ba-item');
			if (it) choose(it);
		});
	}
	function place() {
		var r = btn.getBoundingClientRect(), w = Math.min(380, window.innerWidth - 24);
		panel.style.width = w + 'px';
		var left = Math.min(Math.max(12, r.right - w), window.innerWidth - w - 12);
		panel.style.left = left + 'px';
		panel.style.top = (r.bottom + 8) + 'px';
		panel.style.maxHeight = Math.max(220, window.innerHeight - r.bottom - 24) + 'px';
	}
	function open() {
		if (!panel) build();
		place();
		panel.classList.add('open');
		btn.setAttribute('aria-expanded', 'true');
		input.value = '';
		filter();
		setTimeout(function () { input.focus(); }, 0);
		if (!loaded) load();
	}
	function close() {
		if (!panel || busy) return;
		panel.classList.remove('open');
		btn.setAttribute('aria-expanded', 'false');
	}
	function load() {
		var x = new XMLHttpRequest();
		x.open('GET', 'ba_switch.asp?action=list&_=' + Date.now(), true);
		x.onload = function () {
			var d = null;
			try { d = JSON.parse(x.responseText); } catch (e) { /* bukan JSON */ }
			if (!d || !d.ok) { message(d && d.error ? d.error : 'Daftar BA gagal dimuat.'); return; }
			loaded = true;
			render(d);
		};
		x.onerror = function () { message('Daftar BA gagal dimuat. Periksa koneksi.'); };
		x.send();
	}
	function message(t) { list.innerHTML = ''; list.appendChild(el('div', 'ba-panel-msg', t)); }
	function render(d) {
		list.innerHTML = '';
		items = [];
		if (!d.items.length) { message('Tidak ada BA yang di-assign ke user ini.'); return; }
		var last = null;
		d.items.forEach(function (o) {
			if (o.company !== last) {
				last = o.company;
				var g = el('div', 'ba-group');
				g.appendChild(el('b', '', o.company));
				if (o.companyName) g.appendChild(el('span', '', o.companyName));
				list.appendChild(g);
			}
			var it = el('div', 'ba-item');
			it.setAttribute('role', 'option');
			it.setAttribute('data-code', o.code);
			it.setAttribute('data-name', o.name);
			it.setAttribute('data-search', (o.code + ' ' + o.name + ' ' + o.company + ' ' + (o.companyName || '')).toLowerCase());
			var cur = o.code === d.current;
			if (cur) { it.classList.add('current'); it.setAttribute('aria-selected', 'true'); }
			it.appendChild(el('span', 'ba-code', o.code));
			it.appendChild(el('span', 'ba-name', o.name));
			if (cur) it.appendChild(el('span', 'ba-tag', 'Aktif'));
			list.appendChild(it);
			items.push(it);
		});
		filter();
	}
	function filter() {
		if (!items.length) return;
		var q = input.value.toLowerCase().split(/\s+/).filter(Boolean), shown = 0;
		items.forEach(function (it) {
			var ok = q.every(function (k) { return it.getAttribute('data-search').indexOf(k) !== -1; });
			it.hidden = !ok;
			if (ok) shown++;
		});
		// judul PT disembunyikan bila semua BA di bawahnya tersaring
		Array.prototype.forEach.call(list.querySelectorAll('.ba-group'), function (g) {
			var n = g.nextElementSibling, any = false;
			while (n && !n.classList.contains('ba-group')) { if (!n.hidden) any = true; n = n.nextElementSibling; }
			g.hidden = !any;
		});
		var none = list.querySelector('.ba-none');
		if (!shown && !none) list.appendChild(el('div', 'ba-panel-msg ba-none', 'BA tidak ditemukan.'));
		else if (shown && none) none.parentNode.removeChild(none);
		// saat mengetik, hasil pertama langsung disorot (Enter memilihnya); kode BA yang sama persis didahulukan
		var v = visible(), exact = -1;
		if (q.length === 1) v.forEach(function (it, i) { if (exact < 0 && it.getAttribute('data-code').toLowerCase() === q[0]) exact = i; });
		setActive(q.length ? (exact >= 0 ? exact : 0) : -1);
	}
	function visible() { return items.filter(function (it) { return !it.hidden; }); }
	function setActive(i) {
		var v = visible();
		items.forEach(function (it) { it.classList.remove('active'); });
		active = i;
		if (i >= 0 && v[i]) { v[i].classList.add('active'); v[i].scrollIntoView({ block: 'nearest' }); }
	}
	function keys(e) {
		var v = visible();
		if (e.key === 'ArrowDown') { e.preventDefault(); setActive(Math.min(active + 1, v.length - 1)); }
		else if (e.key === 'ArrowUp') { e.preventDefault(); setActive(Math.max(active - 1, 0)); }
		else if (e.key === 'Enter') { e.preventDefault(); if (v[active]) choose(v[active]); else if (v.length === 1) choose(v[0]); }
		else if (e.key === 'Escape') { close(); btn.focus(); }
	}
	function choose(it) {
		if (busy || it.classList.contains('current')) return;
		var code = it.getAttribute('data-code'), name = it.getAttribute('data-name');
		if (!window.confirm('Ganti BA ke ' + name + ' (' + code + ')?\n\nBA ini juga menjadi BA default saat login berikutnya. ' +
			'Halaman FB-Plants lain yang sedang terbuka perlu dimuat ulang sebelum menyimpan data.')) return;
		busy = true;
		panel.classList.add('busy');
		message('Mengganti BA ke ' + name + '...');
		var x = new XMLHttpRequest();
		x.open('POST', 'ba_switch.asp', true);
		x.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
		x.onload = function () {
			var d = null;
			try { d = JSON.parse(x.responseText); } catch (e) { /* bukan JSON */ }
			if (d && d.ok) { location.reload(); return; }
			busy = false;
			panel.classList.remove('busy');
			window.alert(d && d.error ? d.error : 'Gagal mengganti BA.');
			loaded = false;
			load();
		};
		x.onerror = function () { busy = false; panel.classList.remove('busy'); window.alert('Gagal mengganti BA. Periksa koneksi.'); loaded = false; load(); };
		x.send('action=switch&fcba=' + encodeURIComponent(code) + '&tok=' + encodeURIComponent(btn.getAttribute('data-tok') || ''));
	}

	btn.addEventListener('click', function (e) {
		e.stopPropagation();
		if (panel && panel.classList.contains('open')) close(); else open();
	});
	document.addEventListener('mousedown', function (e) {
		if (panel && panel.classList.contains('open') && !panel.contains(e.target) && !btn.contains(e.target)) close();
	});
	window.addEventListener('resize', function () { if (panel && panel.classList.contains('open')) place(); });
	document.addEventListener('scroll', function () { if (panel && panel.classList.contains('open')) place(); }, true);
})();
