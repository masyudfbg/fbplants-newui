<!-- Tabel laporan ala DataTables + tampilan laporan modern (js/fbp-table-auto.js). Dimuat sekali per halaman. -->
<script type="text/javascript">
if (!window.FbpTableAuto && !window.__fbtLoading) {
	window.__fbtLoading = true;
	var __s = 'scr' + 'ipt', __h = '';
	// dropdown dengan kolom cari (js/fbp-shell.js) bila halaman belum memuatnya sendiri
	if (!document.querySelector('[src*="fbp-shell.js"]')) {
		__h += '<link rel="stylesheet" href="/css/fbp-shell.css?v=20261009d" type="text/css" />' +
			'<' + __s + ' type="text/javascript" src="/js/fbp-shell.js?v=20261016a"></' + __s + '>';
	}
	document.write(__h + '<link rel="stylesheet" href="/css/fbp-table.css?v=20261010c" type="text/css" />' +
		'<' + __s + ' type="text/javascript" src="/js/fbp-table.js?v=20261015a"></' + __s + '>' +
		'<' + __s + ' type="text/javascript" src="/js/fbp-table-auto.js?v=20261016a"></' + __s + '>');
}
</script>
<script>
function tableToExcel(tableID, filename = 'Export.xls') {
    let table = document.getElementById(tableID);

    let html = '<html><head><meta charset="UTF-8"></head><body>' +
               table.outerHTML +
               '</body></html>';

    let blob = new Blob([html], { type: 'application/vnd.ms-excel' });

    let url = URL.createObjectURL(blob);

    let a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();

    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}
</script>

<input type="button" style="float: right; padding: 10px; margin-bottom: 10px; color: white; background-color: green; border-radius: 5px" value="Export to Excel" onclick="tableToExcel('product-table', 'Report.xls')">
