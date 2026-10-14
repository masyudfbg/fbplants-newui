
   <!-- Tabel laporan ala DataTables + tampilan laporan modern (js/fbp-table-auto.js). Dimuat sekali per halaman. -->
   <script type="text/javascript">
   if (!window.FbpTableAuto && !window.__fbtLoading) {
   	window.__fbtLoading = true;
   	var __s = 'scr' + 'ipt', __h = '';
   	// dropdown dengan kolom cari (js/fbp-shell.js) bila halaman belum memuatnya sendiri
   	if (!document.querySelector('[src*="fbp-shell.js"]')) {
   		__h += '<link rel="stylesheet" href="/css/fbp-shell.css?v=20261009d" type="text/css" />' +
   			'<' + __s + ' type="text/javascript" src="/js/fbp-shell.js?v=20261014b"></' + __s + '>';
   	}
   	document.write(__h + '<link rel="stylesheet" href="/css/fbp-table.css?v=20261010c" type="text/css" />' +
   		'<' + __s + ' type="text/javascript" src="/js/fbp-table.js?v=20261009c"></' + __s + '>' +
   		'<' + __s + ' type="text/javascript" src="/js/fbp-table-auto.js?v=20261014b"></' + __s + '>');
   }
   </script>
   <!-- Verifikasi Manager & Ktu -->
   <table width="80%" class="data fbp-sign" align="center">
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr>
 		<td><font face="Verdana" size="1">Approved by</td>
		<td><font face="Verdana" size="1">Prepared by</td>

	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td><font face="Verdana" size="1"><%=Session("FCMANAGER")%></td>
		<td><font face="Verdana" size="1"><%=Session("KTU")%></td>

	   </tr>
   </table>
</body>
</html>