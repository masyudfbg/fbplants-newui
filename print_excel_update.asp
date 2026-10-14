<script type="text/javascript">//<![CDATA[ 

var tableToExcel = (function() {
  var uri = 'data:application/vnd.ms-excel;base64,'
    , template = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head><body><table>{table}</table></body></html>'
    , base64 = function(s) { return window.btoa(unescape(encodeURIComponent(s))) }
    , format = function(s, c) { return s.replace(/{(\w+)}/g, function(m, p) { return c[p]; }) }
  return function(table, name) {
    if (!table.nodeType) table = document.getElementById(table)
	
	/* ################################################################	*/
	/*	Additional	: AST
		Date 		: 28 May 2018
		Function 	: For Filter Column, No Need to Export Excel
	*/
		var cln=table.cloneNode(true);
		var vhide1 = cln.getElementsByClassName('para1');
		
		while(vhide1[0]) {
			vhide1[0].parentNode.removeChild(vhide1[0]);
		}

		var vhide2 = cln.getElementsByClassName('coba');	
		while(vhide2[0]) {
			vhide2[0].parentNode.removeChild(vhide2[0]);
		}
		
		var vhide3 = cln.getElementsByClassName('IgnoreButton');	
		while(vhide3[0]) {
			vhide3[0].parentNode.removeChild(vhide3[0]);
		}
	/* ################################################################	*/
	
    var ctx = {worksheet: name || 'Worksheet', table: cln.innerHTML}
    window.location.href = uri + base64(format(template, ctx))
  }
})()
//]]>  

</script>

 <input class="tableToExcel" onClick="tableToExcel('entry')" value="Export to Excel" type="button"/>

