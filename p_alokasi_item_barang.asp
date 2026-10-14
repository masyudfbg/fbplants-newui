<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml" lang="id">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>FBPlants Reporting - Alokasi Item Barang</title>

<!-- CSS -->
<link rel="stylesheet" href="/css/jquery-ui.css" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
<style>
    body {
    /*background: #ececec;*/
    }
    /*Hidden class for adding and removing*/
    .lds-dual-ring.hidden {
        display: none;
    }

    /*Add an overlay to the entire page blocking any further presses to buttons or other elements.*/
    .overlay {
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100vh;
        background: rgba(0,0,0,.8);
        z-index: 999;
        opacity: 1;
        transition: all 0.5s;
    }
    
    /*Spinner Styles*/
    .lds-dual-ring {
        display: inline-block;
        width: 80px;
        height: 80px;
    }
    .lds-dual-ring:after {
        content: " ";
        display: block;
        width: 64px;
        height: 64px;
        margin: 5% auto;
        border-radius: 50%;
        border: 6px solid #fff;
        border-color: #fff transparent #fff transparent;
        animation: lds-dual-ring 1.2s linear infinite;
    }
    @keyframes lds-dual-ring {
        0% {
            transform: rotate(0deg);
        }
        100% {
            transform: rotate(360deg);
        }
    }
    .huruftebal {
            font-weight: bold;
    }
    .ratakanan {
            text-align: right;
    }

    table {
        font-family: Arial, Helvetica, sans-serif;
        font-size: 12px;
        border-collapse: collapse;
        width: 100%;
    }

    td, th {
        border: 1px solid #ddd;
        padding: 8px;
    }

    tr:nth-child(even){background-color: #f2f2f2;}

    tr:hover {background-color: #ddd;}

    th {
        text-align: left;
        background-color: #2c3e50;
        color: white;
    }

    .dt-button.btn-excel {
        background-color: #2c3e50 !important; /* hijau */
        color: white !important;
        border: none !important;
        border-radius: 6px;
        padding: 5px 12px;
    }

    .dt-button.btn-pdf {
        background-color: #990b19ff !important; /* merah */
        color: white !important;
        border: none !important;
        border-radius: 6px;
        padding: 5px 12px;
    }

    .dataTables_wrapper .dt-buttons {
        margin-bottom: 10px;     /* jarak bawah tombol */
        margin-right: 15px;      /* jarak kanan antar tombol */
    }

    .dataTables_length {
        margin-left: 10px;       /* jarak kiri dropdown tampilkan */
        margin-right: 20px;      /* jarak kanan ke search */
    }

    .dataTables_filter {
        margin-left: 20px;       /* jarak kiri dari dropdown */
    }

    .tab {
        overflow: hidden;
        border: 1px solid #ccc;
        background-color: #f1f1f1;
        margin: 20px;
    }

    /* Style the buttons inside the tab */
    .tab button {
        background-color: inherit;
        float: left;
        border: none;
        outline: none;
        cursor: pointer;
        padding: 14px 16px;
        transition: 0.3s;
        font-size: 12px;
    }

    /* Change background color of buttons on hover */
    .tab button:hover {
        background-color: #ddd;
    }

    /* Create an active/current tablink class */
    .tab button.active {
        background-color: #ccc;
    }

    /* Style the tab content */
    .tabcontent {
        display: none;
        padding: 6px 12px;
        border-top: none;
    }
</style>
<!-- JS -->
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
<script src="/js/x.js"></script>
<script src="/js/xtableheaderfixed.js"></script>
<script>
  $(function() {
    new xTableHeaderFixed('product-table', window);
});


var tableToExcel = (function () {
  var encabezado = '<html><head><meta http-equiv="content-type" content="text/plain; charset=UTF-8"/><style> table, td {border:thin solid black} table {border-collapse:collapse} td {mso-number-format:"\\@";}</style></head><body><table>';
  var dataTable = document.getElementById('producttable').innerHTML
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

</script>
</head>
<%
    pcompany  		= Request("companyitem")
    fcba      		= Request("businessunititem")
    datefrom  		= Request("datefrom")
    dateto    		= Request("dateto")
    groupbarang    	= Request("groupbarang")
%>
<body>
    <table>
        <tr>
            <th colspan="15" style=" background-color:#FFFFFF; color:#000000; border:none">
                    <center>
                        <div id="header-wrapper">
                            <p style="margin-top: 10px">PT. <%=pcompany%><br>
                            Estate: <%=fcba%> | Periode: <%=datefrom & "-" & dateto%><br>
                            Laporan Bulanan Alokasi Item Barang</p>
                        </div>
                    </center>
            </th>
        </tr>
    </table>
    <div class="tab">
        <button class="tablinks" id="defaultOpen" onclick="openTab(event, 'Tab1')">Item Activity</button>
        <button class="tablinks" onclick="openTab(event, 'Tab2')">Item Activity Jobcode</button>
    </div>
    <div id="wrapper" style="margin: -15px;">
        <div id="Tab1" class="tabcontent">
            <!-- Header -->
            <table class="display" id="tablep_alokasi_item_barang1" style="width:100%">
                <thead>
                    <tr>
                        <th rowspan="2">No</th>
                        <th rowspan="2"><center>Company</center></th>
                        <th rowspan="2"><center>FCBA</center></th>
                        <th rowspan="2"><center>Group Name</center></th>
                        <th rowspan="2"><center>Item Code</center></th>
                        <th rowspan="2"><center>Item Name</center></th>
                        <th rowspan="2"><center>Satuan</center></th>
                        <th rowspan="2"><center>Target Code</center></th>
                        <th colspan="2"><center>Realisasi</center></th>
                    </tr>
                    <tr>
                        <th style="background-color: #2c3e50; color: #fff;">Qty</th>
                        <th style="background-color: #2c3e50; color: #fff;">Amount (Rp)</th>
                    </tr>
                </thead>
                <tbody></tbody>
                <tfoot>
                    <tr>
                        <th colspan="7"></th>
                        <th class="text-end">GRAND TOTAL</th>
                        <th></th>
                        <th></th>
                    </tr>
                </tfoot>
            </table>

            <center>
                <div id="loadingIndicator" style="margin: 20px">
                    <span id="loadingjudul">Sedang di proses menampilkan data...</span><br>
                    <img id="loading" src="IMAGES/bartransparent.gif" />
                </div>
            </center>
            
            <input name="pcompany" type="hidden" id="pcompany" value="<%=Request("companyitem")%>" />
            <input name="fcba" type="hidden" id="fcbas" value="<%=Request("businessunititem")%>" />
            <input name="datefrom" type="hidden" id="datefrom" value="<%=Request("datefrom")%>" />
            <input name="dateto" type="hidden" id="dateto" value="<%=Request("dateto")%>" />
            <input name="groupbarang" type="hidden" id="groupbarang" value="<%=Request("groupbarang")%>" />
            <br>
            <!--#include file="Bottom_Report.asp"-->
        </div>
        <div id="Tab2" class="tabcontent">
            <!-- Header -->
            <table class="display" id="tablep_alokasi_item_barang2">
                <thead>
                    <tr>
                        <th rowspan="2"><center>No</center></th>
                        <th rowspan="2"><center>Company</center></th>
                        <th rowspan="2"><center>FCBA</center></th>
                        <th rowspan="2"><center>Job Code</center></th>
                        <th rowspan="2"><center>Activity</center></th>
                        <th rowspan="2"><center>Group Name</center></th>
                        <th rowspan="2"><center>Item Code</center></th>
                        <th rowspan="2"><center>Item Name</center></th>
                        <th rowspan="2"><center>Satuan</center></th>
                        <th rowspan="2"><center>Target Code</center></th>
                        <th colspan="2"><center>Realisasi</center></th>
                    </tr>
                    <tr>
                        <th style="background-color: #2c3e50;color: #fff;">Qty</th>
                        <th style="background-color: #2c3e50;color: #fff;">Amount (Rp)</th>
                    </tr>
                </thead>
                <tbody></tbody>
                <tfoot>
                    <tr>
                        <th colspan="8"></th>
                        <th colspan="2" class="text-end">GRAND TOTAL</th>
                        <th></th>
                        <th></th>
                    </tr>
                </tfoot>
            </table>

            <center>
                <div id="loadingIndicator1" style="margin: 20px">
                    <span id="loadingjudul1">Sedang di proses menampilkan data...</span><br>
                    <img id="loading1" src="IMAGES/bartransparent.gif" />
                </div>
            </center>
            
            <input name="pcompany" type="hidden" id="pcompany" value="<%=Request("companyitem")%>" />
            <input name="fcba" type="hidden" id="fcbas" value="<%=Request("businessunititem")%>" />
            <input name="datefrom" type="hidden" id="datefrom" value="<%=Request("datefrom")%>" />
            <input name="dateto" type="hidden" id="dateto" value="<%=Request("dateto")%>" />
            <input name="groupbarang" type="hidden" id="groupbarang" value="<%=Request("groupbarang")%>" />
            <br>
            <!--#include file="Bottom_Report.asp"-->
        </div>
    </div>
</body>
<script src="//cdn.datatables.net/2.0.8/js/dataTables.min.js"></script>
<link rel="stylesheet" href="https://cdn.datatables.net/1.13.8/css/jquery.dataTables.min.css">
<link rel="stylesheet" href="https://cdn.datatables.net/buttons/2.4.2/css/buttons.dataTables.min.css">
<script src="https://cdn.datatables.net/1.13.8/js/jquery.dataTables.min.js"></script>
<script src="https://cdn.datatables.net/buttons/2.4.2/js/dataTables.buttons.min.js"></script>
<script src="https://cdn.datatables.net/buttons/2.4.2/js/buttons.html5.min.js"></script>
<script src="https://cdn.datatables.net/buttons/2.4.2/js/buttons.print.min.js"></script>

<script src="https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.2.7/pdfmake.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.2.7/vfs_fonts.js"></script>
<!-- FileSaver untuk Excel -->
<script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
<script>
    let tab1Loaded = false;
    let tab2Loaded = false;

    function openTab(evt, tabName) {
        var i, tabcontent, tablinks;
        tabcontent = document.getElementsByClassName("tabcontent");
        for (i = 0; i < tabcontent.length; i++) {
            tabcontent[i].style.display = "none";
        }
        tablinks = document.getElementsByClassName("tablinks");
        for (i = 0; i < tablinks.length; i++) {
            tablinks[i].className = tablinks[i].className.replace(" active", "");
        }
        document.getElementById(tabName).style.display = "block";
        evt.currentTarget.className += " active";

        // --- LOAD DATA HANYA SEKALI ---
        if (tabName == 'Tab1') {
            if (!tab1Loaded) {             // << CEK SUDAH PERNAH LOAD?
                OpenData(43);
                tab1Loaded = true;         // << PERTAMA KALI SAJA
            }
        } 
        else if (tabName == 'Tab2') {
            if (!tab2Loaded) {             // << CEK SUDAH PERNAH LOAD?
                OpenData1(43);
                tab2Loaded = true;
            }
        }
    }


    document.getElementById("defaultOpen").click();

    function OpenData(tabName) {
        let pcompany = $('#pcompany').val();
        let fcba = $('#fcbas').val();
        let datefrom = $('#datefrom').val();
        let dateto = $('#dateto').val();
        let groupbarang = $('#groupbarang').val();

        $('#loadingIndicator').show();

        // Jika DataTable sudah ada, hapus datanya tanpa menghapus header/footer
        if ($.fn.DataTable.isDataTable('#tablep_alokasi_item_barang1')) {
            $('#tablep_alokasi_item_barang1').DataTable().clear().destroy();
        }

        $.post(
            "bkm_viewdata.asp?fcba=" + fcba + 
            "&idview=" + tabName + 
            "&datefrom=" + datefrom + 
            "&dateto=" + dateto + 
            "&companyitem=" + pcompany + 
            "&groupbarang=" + groupbarang + 
            "&querycode=1",
            function (data) {
                try {
                    var json = JSON.parse(data);

                    $('#tablep_alokasi_item_barang1').DataTable({
                        destroy: true,
                        data: json.results,
                        paging: true,
                        pageLength: -1, // tampilkan semua data
                        lengthMenu: [[-1], ["All"]],
                        searching: true,
                        ordering: true,
                        scrollX: true,
                        autoWidth: false,
                        initComplete: function (settings, json) {
                            $('#loadingIndicator').hide();
                        },
                        columns: [
                            {
                                render: function (data, type, row, meta) {
                                    return meta.row + 1;
                                }
                            },
                            { data: "fccompanycode", defaultContent: "" },
                            { data: "fcba", defaultContent: "" },
                            { data: "groupname", defaultContent: "" },
                            { data: "itemcode", defaultContent: "" },
                            { data: "itemname", defaultContent: "" },
                            { data: "satuan", defaultContent: "" },
                            { data: "locationcode", defaultContent: "" }, // Target Code
                            {
                                data: "unit",
                                className: "text-end",
                                defaultContent: "0",
                                render: function (data) {
                                    return data ? parseFloat(data).toLocaleString('id-ID') : '-';
                                }
                            },
                            {
                                data: "amount",
                                className: "text-end",
                                defaultContent: "0",
                                render: function (data) {
                                    if (data == null || data === "") return "-";
                                    return new Intl.NumberFormat('id-ID', {
                                        style: 'currency',
                                        currency: 'IDR',
                                        minimumFractionDigits: 0
                                    }).format(data);
                                }
                            }
                        ],
                        language: {
                            lengthMenu: "Tampilkan _MENU_ baris per halaman",
                            zeroRecords: "Data tidak ditemukan",
                            info: "Menampilkan _TOTAL_ data",
                            infoEmpty: "Tidak ada data",
                            infoFiltered: "(disaring dari total _MAX_ data)",
                            search: "Cari:",
                            paginate: {
                                first: "Awal",
                                last: "Akhir",
                                next: "Berikutnya",
                                previous: "Sebelumnya"
                            }
                        },
                        dom: 'Blfrtip',
                        buttons: [
                            {
                                extend: 'excelHtml5',
                                text: 'Export Excel',
                                className: 'btn-excel',
                                title: 'Monitoring Stock Harian'
                            },
                            {
                                extend: 'pdfHtml5',
                                text: 'Export PDF',
                                className: 'btn-pdf',
                                orientation: 'landscape',
                                pageSize: 'A4',
                                title: 'Monitoring Stock Harian'
                            }
                        ],

                        // Footer Grand Total
                        footerCallback: function (row, data, start, end, display) {
                            var api = this.api();

                            // Konversi ke number
                            var intVal = function (i) {
                                if (typeof i === 'string') {
                                    return i.replace(/[^0-9,-]+/g, '').replace(',', '.') * 1 || 0;
                                } else if (typeof i === 'number') {
                                    return i;
                                } else {
                                    return 0;
                                }
                            };

                            // Total Qty (kolom ke-8)
                            var totalQty = api
                                .column(8, { page: 'current' })
                                .data()
                                .reduce(function (a, b) {
                                    return intVal(a) + intVal(b);
                                }, 0) || 0;

                            // Total Amount (kolom ke-9)
                            var totalAmount = api
                                .column(9, { page: 'current' })
                                .data()
                                .reduce(function (a, b) {
                                    return intVal(a) + intVal(b);
                                }, 0) || 0;

                            // Format rupiah
                            var totalAmountFormatted = new Intl.NumberFormat('id-ID', {
                                style: 'currency',
                                currency: 'IDR',
                                minimumFractionDigits: 0
                            }).format(totalAmount);

                            // Tulis ke footer
                            $(api.column(7).footer()).html('<b>GRAND TOTAL</b>');
                            $(api.column(8).footer()).html('<b>' + totalQty.toLocaleString('id-ID') + '</b>');
                            $(api.column(9).footer()).html('<b>' + totalAmountFormatted + '</b>');
                        }
                    });

                    $('#loadingIndicator').hide();

                } catch (e) {
                    console.error("JSON parse error:", e, data);
                    $('#loadingIndicator').hide();
                }
            }
        );
    }


    function OpenData1(tabName){
        let pcompany  		= $('#pcompany').val();
        let fcba      		= $('#fcbas').val();
        let datefrom  		= $('#datefrom').val();
        let dateto    		= $('#dateto').val();
        let groupbarang    	= $('#groupbarang').val();
        $('#loadingIndicator1').show();

        $.post("bkm_viewdata.asp?fcba=" + fcba + "&idview=" + tabName  + "&datefrom=" + datefrom + "&dateto=" + dateto + '&companyitem=' + pcompany + '&groupbarang=' + groupbarang + '&querycode=2', 
            function(data) {
                try {
                    var json  = JSON.parse(data);

                    $('#tablep_alokasi_item_barang2').DataTable({
                        destroy: true,
                        data: json.results,
                        paging: false,
                        searching: true,
                        ordering: true,
                        scrollY: "400px",
                        pageLength: -1,
                        scrollX: true,
                        autoWidth: false,
                        initComplete: function(settings, json) {
                            // Hide loading indicator after initialization
                            $('#loadingIndicator1').hide(); 
                        },
                        columns: [
                            { 
                                render: function (data, type, row, meta) {
                                    return meta.row + meta.settings._iDisplayStart + 1;
                                }
                            },
                            { data: "fccompanycode", defaultContent: "" },
                            { data: "fcba", defaultContent: "" },
                            { data: "jobcode", defaultContent: "" },
                            { data: "jobname", defaultContent: "" },
                            { data: "groupname", defaultContent: "" },
                            { data: "itemcode", defaultContent: "" },
                            { data: "itemname", defaultContent: "" },
                            { data: "satuan", defaultContent: "" },
                            { data: "locationcode", defaultContent: "" },
                            { data: "unit", className: "text-end", defaultContent: "0" },
                            { 
                                data: "amount",
                                defaultContent: "0",
                                render: function (data, type, row) {
                                if (data == null || data === "") return "-";
                                // format rupiah
                                return new Intl.NumberFormat('id-ID', { 
                                    style: 'currency', 
                                    currency: 'IDR', 
                                    minimumFractionDigits: 0 
                                }).format(data);
                                },
                                className: "text-end"
                            }
                        ],
                        language: {
                            lengthMenu: "Tampilkan _MENU_ baris per halaman",
                            zeroRecords: "Data tidak ditemukan",
                            info: "Menampilkan _START_ sampai _END_ dari _TOTAL_ data",
                            infoEmpty: "Tidak ada data",
                            infoFiltered: "(disaring dari total _MAX_ data)",
                            search: "Cari:",
                            paginate: {
                                first: "Awal",
                                last: "Akhir",
                                next: "Berikutnya",
                                previous: "Sebelumnya"
                            }
                        },
                        dom: 'Blfrtip',
                        lengthMenu: [[-1], ["All"]],
                        buttons: [
                            {
                                extend: 'excelHtml5',
                                text: 'Export Excel',
                                className: 'btn-excel',
                                title: 'Item Activity Jobcode'
                            },
                            {
                                extend: 'pdfHtml5',
                                text: 'Export PDF',
                                className: 'btn-pdf',
                                orientation: 'landscape',
                                pageSize: 'A4',
                                title: 'Item Activity Jobcode'
                            }
                        ],

                        footerCallback: function (row, data, start, end, display) {
                            var api = this.api();

                            // Fungsi untuk konversi angka dari format string
                            var intVal = function (i) {
                                return typeof i === 'string'
                                    ? i.replace(/[^0-9,-]+/g, '').replace(',', '.') * 1
                                    : typeof i === 'number'
                                    ? i
                                    : 0;
                            };

                            // Hitung total Qty (kolom ke-10 -> index 10)
                            var totalQty = api
                                .column(10, { page: 'current' })
                                .data()
                                .reduce(function (a, b) {
                                    return intVal(a) + intVal(b);
                                }, 0);

                            // Hitung total Amount (kolom ke-11 -> index 11)
                            var totalAmount = api
                                .column(11, { page: 'current' })
                                .data()
                                .reduce(function (a, b) {
                                    return intVal(a) + intVal(b);
                                }, 0);

                            // Format ke rupiah
                            var totalAmountFormatted = new Intl.NumberFormat('id-ID', { 
                                style: 'currency', 
                                currency: 'IDR', 
                                minimumFractionDigits: 0 
                            }).format(totalAmount);

                            // Tampilkan di footer
                            $(api.column(9).footer()).html('<b>GRAND TOTAL</b>');
                            $(api.column(10).footer()).html('<b>' + totalQty.toLocaleString('id-ID') + '</b>');
                            $(api.column(11).footer()).html('<b>' + totalAmountFormatted + '</b>');
                        }
                    });
                    
                    $('#loadingIndicator1').hide();
                    // $('#loadingjudul').hide();
                    // $('#loading').hide();
                } catch (e) {
                    console.error("JSON parse error:", e, data);
                    return;
                }
            }
        );
    }

    function numformat(val) {
		var nStr = val.toString();
		nStr = nStr.replace(/\,/g, "");
		var x = nStr.split('.');
		var x1 = x[0];
		var x2 = x.length > 1 ? '.' + x[1] : '';
		var rgx = /(\d+)(\d{3})/;
		while (rgx.test(x1)) {
			x1 = x1.replace(rgx, '$1' + '.' + '$2');
		}
		return x1 + x2;
	}
</script>
</html>
