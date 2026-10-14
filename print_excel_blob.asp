<script type="text/javascript">
//<![CDATA[ 

var tableToExcel = (function () {
  var encabezado = '<html><head><meta http-equiv="content-type" content="text/plain; charset=UTF-8"/><style> table, td {border:thin solid black} table {border-collapse:collapse} td {mso-number-format:"\\@";}</style></head><body><table>';
  var dataTable = document.getElementById('product-table').innerHTML
  var piePagina = "</table></body></html>";
  var tabla = encabezado + dataTable + piePagina;
  var myBlob =  new Blob( [tabla] , {type:'text/html'});
  var url = window.URL.createObjectURL(myBlob);
  var a = document.createElement("a");
  document.body.appendChild(a);
  a.href = url;
  a.download = "export_excel.xls";
  a.click();
  setTimeout(function() {window.URL.revokeObjectURL(url);},0);
});

//]]>  
</script>


 <input class="tableToExcel" onClick="tableToExcel()" value="Export to Excel" type="button">

