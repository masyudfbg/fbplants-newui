<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->

<%
' ============================================================
' HELPER FUNCTIONS - harus di luar blok eksekusi utama
' ============================================================
Function safeNum(val)
    If IsNull(val) Or IsEmpty(val) Then
        safeNum = 0
    Else
        safeNum = CDbl(val)
    End If
End Function

Function safeDiv(numerator, denominator)
    If denominator = 0 Then
        safeDiv = 0
    Else
        safeDiv = numerator / denominator
    End If
End Function

Function fmtCell(val)
    Dim n
    n = safeNum(val)
    If n = 0 Then
        fmtCell = "<span class='zero-val'>-</span>"
    Else
        fmtCell = formatNumber(n, 0)
    End If
End Function

Function fmtTotal(val)
    If val = 0 Then
        fmtTotal = "-"
    Else
        fmtTotal = formatNumber(val, 0)
    End If
End Function
%>

<html>
	<head>
		<title>FBPlants - HPP BIBIT VIEW</title>
		<link href="/js/tabbedpanels/SpryTabbedPanels.css" rel="stylesheet" type="text/css"/>
		<script src="/js/tabbedpanels/SpryTabbedPanels.js" type="text/javascript"></script>
		<script src="/js/jquery-1.9.1.js"></script>
		<script src="/js/jquery-ui.js"></script>
		<link rel="stylesheet" href="css/jquery-ui.css" />
		<link rel="StyleSheet" href="style-content.css" type="text/css" />
		<link href="/js/spry/widgets/textfieldvalidation/SpryValidationTextField.css" rel="stylesheet" type="text/css"/>
		<script src="/js/spry/widgets/textfieldvalidation/SpryValidationTextField.js" type="text/javascript"></script>
		<link rel="stylesheet" href="https://cdn.datatables.net/1.13.8/css/jquery.dataTables.min.css">
		<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
		<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
	</head>
	<style>
        *, *::before, *::after { box-sizing: border-box; }

		#wrapper {
            padding: 24px 20px;
        }

        /* ============ TABLE CARD ============ */
        .table-card {
            background: #ffffff;
            border-radius: 12px;
            box-shadow: 0 2px 16px rgba(0,0,0,0.08), 0 0 0 1px rgba(255,168,0,0.12);
            overflow: hidden;
        }

        .table-card-header {
            display: flex;
            align-items: center;
            gap: 10px;
            padding: 14px 20px;
            background: linear-gradient(135deg, #0c5b24 0%, #0c5b24 100%);
            color: #fff;
        }

        .table-card-header i {
            font-size: 15px;
        }

        .table-card-header h2 {
            font-size: 13px;
            font-weight: 700;
            letter-spacing: 0.6px;
            text-transform: uppercase;
            margin: 0;
        }

        /* ============ TABLE SCROLL ============ */
        .table-scroll {
            overflow-x: auto;
        }

        /* ============ STICKY HEADER + SCROLLABLE BODY ============ */
        .table-container {
            position: relative;
            max-height: 420px;
            overflow-y: auto;
            overflow-x: auto;
        }

        /* Custom scrollbar */
        .table-container::-webkit-scrollbar {
            width: 6px;
            height: 6px;
        }
        .table-container::-webkit-scrollbar-track {
            background: #f0f1f5;
        }
        .table-container::-webkit-scrollbar-thumb {
            background: #FFA800;
            border-radius: 4px;
        }
        .table-container::-webkit-scrollbar-thumb:hover {
            background: #FF8C00;
        }

        /* ============ MAIN TABLE ============ */
        table#table2 {
            width: 100%;
            border-collapse: collapse;
            font-size: 12px;
            font-family: 'Plus Jakarta Sans', sans-serif;
        }

        /* HEADER - sticky */
        table#table2 thead tr th {
            background: #1e2235;
            color: #FFA800;
            font-size: 11px;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            padding: 10px 12px;
            white-space: nowrap;
            border: none;
            text-align: center;
            position: sticky;
            top: 0;
            z-index: 5;
        }

        /* BODY ROWS */
        table#table2 tbody tr {
            border-bottom: 1px solid #000;
            transition: background 0.12s ease;
        }

        table#table2 tbody tr:nth-child(even) {
            background: #fafbff;
        }

        table#table2 tbody tr:hover {
            background: #fff8e6;
        }

        table#table2 tbody td {
            padding: 8px 12px;
            font-size: 12px;
            color: #2d3142;
            border: 1px solid #000;
            vertical-align: middle;
        }

        /* NO column */
        table#table2 tbody td:first-child {
            color: #9399b2;
            font-weight: 600;
            text-align: center;
            font-size: 11px;
        }

        /* FCBA */
        table#table2 tbody td:nth-child(2) {
            font-weight: 600;
            color: #1e2235;
        }

        /* KODE BATCH - mono */
        table#table2 tbody td:nth-child(3) {
            font-family: 'JetBrains Mono', monospace;
            font-size: 11px;
            color: #5c6080;
        }

        /* TANGGAL */
        table#table2 tbody td:nth-child(4) {
            font-size: 11px;
            color: #5c6080;
            white-space: nowrap;
        }

        /* Numeric columns */
        table#table2 tbody td:nth-child(n+5) {
            text-align: right;
            font-variant-numeric: tabular-nums;
            font-family: 'JetBrains Mono', monospace;
            font-size: 11px;
        }

        /* Zero / dash */
        .zero-val {
            color: #d0d3e0;
        }

        /* FOOTER TABLE - always visible below scroll area */
        table#table2 tfoot tr {
            background: linear-gradient(90deg, #1e2235 0%, #2d3557 100%);
        }

        table#table2 tfoot td {
            padding: 10px 12px;
            font-size: 11px;
            font-weight: 700;
            color: #FFA800;
            text-align: right;
            font-family: 'JetBrains Mono', monospace;
            font-variant-numeric: tabular-nums;
            border: none;
        }

        table#table2 tfoot td.footer-label {
            text-align: left;
            font-family: 'Plus Jakarta Sans', sans-serif;
            font-size: 11px;
            letter-spacing: 0.7px;
            text-transform: uppercase;
            color: #fff;
        }

        table#table2 tfoot td.footer-label span {
            color: #FFA800;
        }

        /* FOOTER FIXED TABLE (outside scroll) */
        .footer-fixed-wrap {
            overflow-x: auto;
            border-top: 1px solid #000;
        }

        table#table-footer {
            width: 100%;
            border-collapse: collapse;
            font-size: 12px;
            table-layout: fixed;
        }

        table#table-footer td {
            padding: 10px 12px;
            font-size: 11px;
            font-weight: 700;
            color: #FFA800;
            text-align: right;
            font-family: 'JetBrains Mono', monospace;
            font-variant-numeric: tabular-nums;
            border: 1px solid #000;
            background: linear-gradient(90deg, #1e2235 0%, #2d3557 100%);
            white-space: nowrap;
        }

        table#table-footer td.footer-label {
            text-align: left;
            font-family: 'Plus Jakarta Sans', sans-serif;
            letter-spacing: 0.7px;
            text-transform: uppercase;
            color: #fff;
        }

        table#table-footer td.footer-label span {
            color: #FFA800;
        }

        /* ============ HELPERS ============ */
        .ratakiri   { text-align: left !important; }
        .ratakanan  { text-align: right !important; }
        .ratatengah { text-align: center !important; }
        .tebal      { font-weight: 700; }
	</style>
<body>
<link rel="stylesheet" href="2col_leftNav.css" type="text/css" style="masthead" />
<script type="text/javascript">
		function cancelFullScreen(el) {
            var requestMethod = el.cancelFullScreen||el.webkitCancelFullScreen||el.mozCancelFullScreen||el.exitFullscreen;
            if (requestMethod) { // cancel full screen.
                requestMethod.call(el);
            } else if (typeof window.ActiveXObject !== "undefined") { // Older IE.
                var wscript = new ActiveXObject("WScript.Shell");
                if (wscript !== null) {
                    wscript.SendKeys("{F11}");
                }
            }
        }

        function requestFullScreen(el) {
            // Supports most browsers and their versions.
            var requestMethod = el.requestFullScreen || el.webkitRequestFullScreen || el.mozRequestFullScreen || el.msRequestFullScreen;

            if (requestMethod) { // Native full screen.
                requestMethod.call(el);
            } else if (typeof window.ActiveXObject !== "undefined") { // Older IE.
                var wscript = new ActiveXObject("WScript.Shell");
                if (wscript !== null) {
                    wscript.SendKeys("{F11}");
                }
            }
            return false
        }

        function toggleFull() {
            var elem = document.body; // Make the body go full screen.
            var isInFullScreen = (document.fullScreenElement && document.fullScreenElement !== null) ||  (document.mozFullScreen || document.webkitIsFullScreen);

            if (isInFullScreen) {
                cancelFullScreen(document);
            } else {
                requestFullScreen(elem);
				 alert("Tekan 'F11' pada Keyboard Anda, untuk selalu Full Screen.");
            }
            return false;
        }

		
</script>
<style>
#siteInfo{ background-image:url(IMAGES/login_02.gif); background-repeat: repeat-x; height:107px; border:none; left: 0;
    position: fixed;
    right: 0;
    top: 0; }
.colR{ float:right; margin-top:30px;  }
.logo { margin-top:10px; width:30%; float:left; }
</style>
<div id="siteInfo" style="height:90px; border-bottom:4px #FFCC00 solid;">
    <img src="/css/login/images/fbg-logo.png" alt="FBG" class="fbp-logo" height="44" style="margin-left:10px; float:left;">
	<!--<font color="yellow" pointsize="7px">iPlaS</font>-->
	&nbsp;&nbsp;&nbsp;
	<h1 style="float:left; width:300px; margin-top:21px; margin-left:10px;"><font color="#FFCC00" style="margin-top:5px;">First Borneo Plantation System</font></h1>
	<div id="globalNav">
       <%if Session("USERNAME")="" then
		Response.Redirect("logout.asp")
		%>
		<a class="mn_top" href="logout.asp"><span>log-in</span></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<B>Please login first...</B><a class="mn_top"  href="login.Asp"><i class="home_icon">&nbsp;&nbsp;&nbsp;&nbsp;</i><span>Home</span></a></div>
		<%else%>
		    <a class="mn_top" style="width: 70px;text-align:center" href="logout.asp"><i class="off_icon">&nbsp;&nbsp;&nbsp;&nbsp;</i><span>log-out</span></a>
            <a class="mn_top" style="width: 66px;text-align:center"  href="login.Asp"><i class="home_icon">&nbsp;&nbsp;&nbsp;&nbsp;</i><span>Home</span></a>
			<a id="button" class="mn_top" style="width: 91px;text-align:center" href="javascript:void(0);" onclick="toggleFull();">
				<i class="full">&nbsp;&nbsp;&nbsp;&nbsp;</i><span>Full Screen</span>
			</a>
            <!-- 
			<a id="button" class="mn_top" target="_blank" href="http://103.108.126.134:8080/IPLAS/index.php/" >
				<i class="home_icon">&nbsp;&nbsp;&nbsp;&nbsp;</i><span>Plus +</span>
			</a>
            -->
		
            <a id="button" class="mn_top" style="width: 40px;text-align:center" href="http://firstborneo.com/dx/"><span>DX</span></a>

		</div>
      
        <br/>
        <div class="identity" ><b><%=session("FULLNAME")%></b> <font size="4" color="#FFCC00">|</font> <b>BA : <%=Session("BA")%> ( <%=session("fcBA")%> )</b> <font size="4"  color="#FFCC00">|</font> Current Period : <%=FormatDateTime(Cdate(Now),1)%>
		<%end if%> </div>
	
</div>

<br>
<br>
<br>
<br>
<br>

<script language="JavaScript" src='/js/x.js'></script>
<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
<script language="JavaScript">
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window);
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);
</script>
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	newwin = window.open(winurl,winname,winfeatures);
}
</script>

<div id="wrapper">
    <form name="frmEntry">
        <table>
            <tr>
                <td>Batch No</td>
                <td>:</td>
                <td><%=Request("batch")%></td>
            </tr>
            <tr>
                <td>Periode Tanggal</td>
                <td>:</td>
                <td><%=Request("from")%> s/d <%=Request("to")%></td>
            </tr>
            <tr>
                <td>Business Unit</td>
                <td>:</td>
                <td><%=Request("fcba")%></td>
            </tr>
        </table>
        <br>
        <div class="table-card">
            <div class="table-card-header">
                <h2 style="color: #fff">HPP Bibit View</h2>
            </div>

            <div class="table-container">
                <table id="table2">
                    <thead>
                        <tr id="trheader2">
                            <th>NO</th>
                            <th>FCBA</th>
                            <th>KODE BATCH</th>
                            <th>TANGGAL</th>
                            <th>KECAMBAH</th>
                            <th>TANAM KECAMBAH</th>
                            <th>PRE NURSERY (Stock)</th>
                            <th>MAIN NURSERY (Stock)</th>
                            <th>TRANSFER</th>
                            <th>ALOKASI</th>
                            <th>PERSIAPAN</th>
                            <th>PRE NURSERY (Rp)</th>
                            <th>MAIN NURSERY (Rp)</th>
                            <th>TOTAL (Rp)</th>
                            <th>HPP (Rp)</th>
                        </tr>
                    </thead>
                    <tbody>

                    <%
                    On Error Resume Next

                    Dim sqldtlupah, rsdtlupah, no
                    Dim company47, fcba47, tahun47, bulan47, from47, to47, batch47
                    Dim whereMonth, whereDate, wherebatch, wheresql47

                    ' Totals per column
                    Dim tot_kecambah, tot_tanam, tot_pre_stock, tot_main_stock
                    Dim tot_transfer, tot_alokasi, tot_persiapan, tot_pre_rp, tot_main_rp, tot_total_rp
                    tot_kecambah   = 0
                    tot_tanam      = 0
                    tot_pre_stock  = 0
                    tot_main_stock = 0
                    tot_transfer   = 0
                    tot_alokasi    = 0
                    tot_persiapan  = 0
                    tot_pre_rp     = 0
                    tot_main_rp    = 0
                    tot_total_rp   = 0

                    company47 = Request("company")
                    fcba47    = Request("fcba")
                    tahun47   = Request("tahun")
                    bulan47   = Request("month")
                    from47    = Request("from")
                    to47      = Request("to")
                    batch47   = Request("batch")

                    If trim(from47) = "0" Then
                        whereDate = ""
                    Else
                        whereDate = " AND a.DOCDATE BETWEEN to_date('" & from47 & "', 'mm/dd/yyyy') AND to_date('" & to47 & "', 'mm/dd/yyyy') AND a.ACCOUNTINGPERIOD <= 12 "
                    End If

                    If trim(batch47) <> "-" Then
                        wherebatch = " AND COALESCE(n.NURSERYCODE, b.LOCATIONCODE) = '" & batch47 & "' "
                    Else
                        wherebatch = " "
                    End If

                    If trim(fcba47) = "-" Then
                        wheresql47 = " a.fcba IN ( " & _
                                    "   SELECT fccode FROM businessunit " & _
                                    "   WHERE fccompanycode = '" & company47 & "' " & _
                                    "   AND isactive = 'TRUE' AND fctype='E' " & _
                                    " ) "
                    Else
                        wheresql47 = " a.fcba ='" & fcba47 & "' "
                    End If

                        sqldtlupah = "WITH jobmaster_group AS (" & _
                                    "    SELECT fccode," & _
                                    "        jobgroupcode," & _
                                    "        MAX(CASE WHEN jobgroupcode = '1310200' THEN 1 ELSE 0 END) AS is_group_1310200," & _
                                    "        MAX(CASE WHEN jobgroupcode = '1310300' THEN 1 ELSE 0 END) AS is_group_1310300" & _
                                    "    FROM jobmaster" & _
                                    "    WHERE jobgroupcode IN ('1310200', '1310300')" & _
                                    "    GROUP BY fccode, jobgroupcode" & _
                                    ")" & _
                                    "SELECT b.FCBA," & _
                                    "    b.LOCATIONTYPE," & _
                                    "    b.LOCATIONCODE," & _
                                    "    n.DATEPLANTED," & _
                                    "    (b.AMOUNTDEBIT + b.AMOUNTKREDIT) AS TOTAL," & _
                                    "    COALESCE(n.NURSERYCODE, b.LOCATIONCODE) AS NURSERYCODE," & _
                                    "    c.FCNAME," & _
                                    "    CASE WHEN b.JOBCODE = '1310203' THEN b.UNIT ELSE 0 END AS KECAMBAH," & _
                                    "    CASE WHEN b.JOBCODE = '1310205' THEN b.UNIT ELSE 0 END AS TANAM_KECAMBAH," & _
                                    "    CASE WHEN b.JOBCODE = '1310214' THEN b.UNIT ELSE 0 END AS PRE_NURSERY," & _
                                    "    CASE WHEN b.JOBCODE = '1310305' THEN b.UNIT ELSE 0 END AS MAIN_NURSERY," & _
                                    "    CASE WHEN b.JOBCODE = '1310401' THEN b.UNIT ELSE 0 END AS TRANSFER," & _
                                    "    CASE WHEN b.JOBCODE = '1310401' THEN b.AMOUNTDEBIT + b.AMOUNTKREDIT ELSE 0 END AS ALOKASI," & _
                                    "    CASE WHEN jg.is_group_1310200 = 1 THEN b.AMOUNTDEBIT + b.AMOUNTKREDIT ELSE 0 END AS PRENURSERY2," & _
                                    "    CASE WHEN jg.is_group_1310300 = 1 THEN b.AMOUNTDEBIT + b.AMOUNTKREDIT ELSE 0 END AS MAIN_NURSERY2," & _
                                    "    0 AS SALDOAWAL," & _
                                    "    b.AMOUNTDEBIT AS DEBET," & _
                                    "    b.AMOUNTKREDIT AS KREDIT," & _
                                    "    0 AS SALDOAKHIR," & _
                                    "    0 AS STOK," & _
                                    "    0 AS HPP," & _
                                    "    n.DATEPLANTED AS NURSERY_DATEPLANTED," & _
                                    "    n.DESCRIPTION" & _
                                    " FROM BKMJOURNALHDR a" & _
                                    " LEFT JOIN BKMJOURNALDTL b ON a.DOCNO = b.DOCNO" & _
                                    " LEFT JOIN JOBMASTER c ON b.JOBCODE = c.FCCODE" & _
                                    " LEFT JOIN jobmaster_group jg ON b.JOBCODE = jg.fccode" & _
                                    " LEFT JOIN (" & _
                                    "     SELECT NURSERYCODE, DATEPLANTED, DESCRIPTION" & _
                                    "     FROM (" & _
                                    "         SELECT" & _
                                    "             NURSERYCODE," & _
                                    "             DATEPLANTED," & _
                                    "             DESCRIPTION," & _
                                    "             ROW_NUMBER() OVER (" & _
                                    "                 PARTITION BY NURSERYCODE" & _
                                    "                 ORDER BY DATEPLANTED DESC NULLS LAST" & _
                                    "             ) AS RN" & _
                                    "         FROM NURSERY" & _
                                    "     ) x" & _
                                    "     WHERE RN = 1" & _
                                    " ) n ON b.LOCATIONCODE = n.NURSERYCODE " &_
                                    " WHERE " & wheresql47 & wherebatch & whereMonth & _
                                    "   AND c.FCCODE IN ('1310203', '1310205', '1310214', '1310305', '1310401')" & _
                                    whereDate & _
                                    wherebatch & _
                                    " ORDER BY a.DOCDATE ASC"

                    ' response.write sqldtlupah
                    Set rsdtlupah = connStr.Execute(sqldtlupah)

                    If Err.Number <> 0 Then
                        Response.Write "<tr><td colspan='15' style='color:red;padding:12px;'>Query Error: " & Err.Description & "</td></tr>"
                    Else
                        no = 1
                        Do While Not rsdtlupah.EOF

                            ' Hitung per baris — HARUS di dalam loop, setelah recordset ada
                            Dim row_total_rp, row_hpp, row_main_stock_val
                            row_main_stock_val = safeNum(rsdtlupah("main_nursery"))
                            row_total_rp       = safeNum(rsdtlupah("total"))
                            row_hpp            = safeDiv(row_total_rp, row_main_stock_val)

                            ' Akumulasi total
                            tot_kecambah   = tot_kecambah   + safeNum(rsdtlupah("kecambah"))
                            tot_tanam      = tot_tanam      + safeNum(rsdtlupah("tanam_kecambah"))
                            tot_pre_stock  = tot_pre_stock  + safeNum(rsdtlupah("pre_nursery"))
                            tot_main_stock = tot_main_stock + row_main_stock_val
                            tot_transfer   = tot_transfer   + safeNum(rsdtlupah("transfer"))
                            tot_alokasi    = tot_alokasi    + safeNum(rsdtlupah("alokasi"))
                            tot_persiapan  = tot_persiapan  + safeNum(rsdtlupah("persiapan"))
                            tot_pre_rp     = tot_pre_rp     + safeNum(rsdtlupah("prenursery2"))
                            tot_main_rp    = tot_main_rp    + safeNum(rsdtlupah("main_nursery2"))
                            tot_total_rp   = tot_total_rp   + row_total_rp
                    %>
                            <tr id="tru<%=no%>">
                                <td><%=no%></td>
                                <td class="ratakiri"><%=rsdtlupah("fcba")%></td>
                                <td class="ratakiri"><%=rsdtlupah("locationcode")%></td>
                                <td class="ratatengah"><%=rsdtlupah("docdate")%></td>
                                <td><%=fmtCell(rsdtlupah("kecambah"))%></td>
                                <td><%=fmtCell(rsdtlupah("tanam_kecambah"))%></td>
                                <td><%=fmtCell(rsdtlupah("pre_nursery"))%></td>
                                <td><%=fmtCell(rsdtlupah("main_nursery"))%></td>
                                <td><%=fmtCell(rsdtlupah("transfer"))%></td>
                                <td><%=fmtCell(rsdtlupah("alokasi"))%></td>
                                <td><%=fmtCell(rsdtlupah("persiapan"))%></td>
                                <td><%=fmtCell(rsdtlupah("prenursery2"))%></td>
                                <td><%=fmtCell(rsdtlupah("main_nursery2"))%></td>
                                <td><%=fmtCell(row_total_rp)%></td>
                                <td><%=fmtCell(row_hpp)%></td>
                            </tr>
                    <%
                            no = no + 1
                            rsdtlupah.MoveNext
                        Loop
                    End If

                    rsdtlupah.Close
                    Set rsdtlupah = Nothing
                    %>

                    </tbody>
                    <tfoot>
                        <tr>
                            <td colspan="4" class="footer-label">
                                Total &nbsp;<span>(<%=no-1%> baris)</span>
                            </td>
                            <td><%=fmtTotal(tot_kecambah)%></td>
                            <td><%=fmtTotal(tot_tanam)%></td>
                            <td><%=fmtTotal(tot_pre_stock)%></td>
                            <td><%=fmtTotal(tot_main_stock)%></td>
                            <td><%=fmtTotal(tot_transfer)%></td>
                            <td><%=fmtTotal(tot_alokasi)%></td>
                            <td><%=fmtTotal(tot_persiapan)%></td>
                            <td><%=fmtTotal(tot_pre_rp)%></td>
                            <td><%=fmtTotal(tot_main_rp)%></td>
                            <td><%=fmtTotal(tot_total_rp)%></td>
                            <td><%=fmtTotal(safeDiv(tot_total_rp, tot_main_stock))%></td>
                        </tr>
                    </tfoot>
                </table>
            </div><!-- /table-container -->
        </div><!-- /table-card -->

    </form>
</div>

</body>
<!--#include file="bottom.jsp"-->
</html>
<link href="/js/select2/select2.min.css" rel="stylesheet" type="text/css"/>
<script src="/js/select2/select2.min.js" type="text/javascript"></script>
<script src="/distexcel/tableToExcel.js"></script>
<script src="/distexcel/jquery.tableToExcel.js"></script>

<script type="text/javascript">
function actionBtn() {
    var bkmfrom = $("#bkmfrom").val();
    var bkmto   = $("#bkmto").val();

    if (bkmfrom != "") {
        $('#loadingjudul').show();
        $('#loading').show();
        $.post("bkm_viewdata.asp?bkmfrom=" + bkmfrom + "&bkmto=" + bkmto, function(data) {
            $('#bodytable').empty();
            var json     = JSON.parse(data);
            var jumlview = json.results.length;
            var no = 1;
            for (var i = 0; i < jumlview; i++) {
                var docketlist = '';
                docketlist += '<tr id="tru' + i + '">';
                docketlist += '<td align="left">'   + json.results[i].bkmno + '</td>';
                docketlist += '<td align="left">'   + json.results[i].createdate + '</td>';
                docketlist += '<td align="right">'  + numformat(json.results[i].totalbkmupah) + '</td>';
                docketlist += '<td align="right">'  + json.results[i].totalbkmmaterial + '</td>';
                docketlist += '<td align="right">'  + json.results[i].totalbkmvehicle + '</td>';
                docketlist += '<td align="right">'  + json.results[i].totalbkmwip + '</td>';
                docketlist += '<td align="center"><a target="_blank" href="bkm_operasional_edit.Asp?bkmno=' + json.results[i].bkmno + '"><img src="IMAGES/edit.gif" alt="Edit Data" border="0"></a></td>';
                docketlist += '<td align="center"><a target="_blank" href="bkm_jurnal_input.Asp?tglbkm='   + json.results[i].createdate + '"><img src="IMAGES/view.png" alt="View Data" border="0"></a></td>';
                docketlist += '</tr>';
                $('#bodytable').append(docketlist);
                no++;
            }
            $("#jumldata").val(jumlview);
            $('#loadingjudul').hide();
            $('#loading').hide();
        });
    } else {
        $('#bodytable').empty();
    }
}
</script>
