<script type="text/javascript">//<![CDATA[ 

function tableToExcel(tableID, filename = 'export.xls') {
    var table = document.getElementById(tableID);
    var html = '<html xmlns:o="urn:schemas-microsoft-com:office:office" ' +
               'xmlns:x="urn:schemas-microsoft-com:office:excel" ' +
               'xmlns="https://www.w3.org/TR/html401/">' +
               '<head><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet>' +
               '<x:Name>Sheet1</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions>' +
               '</x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->' +
               '</head><body><table>' + table.innerHTML + '</table></body></html>';

    var blob = new Blob([html], { type: 'application/vnd.ms-excel' });
    var url = URL.createObjectURL(blob);
    
    var a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    
    setTimeout(() => {
        URL.revokeObjectURL(url);
        document.body.removeChild(a);
    }, 100);
}


</script>

<input class="tableToExcel" onClick="tableToExcel('product-table')" value="Export to Excel" type="button">

