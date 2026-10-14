<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="UTF-8">
<title>Monitoring Stock Pupuk &amp; Agrochemical</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
    :root {
        --brand: #1B7863;
        --brand-dark: #12584a;
        --brand-tint: #e8f3f0;
        --accent: #F7BD28;
        --ink: #17262a;
        --muted: #5d6d69;
        --line: #e1e7e5;
        --bg: #f3f6f5;
    }
    * { box-sizing: border-box; }
    body {
        margin: 0;
        padding: 20px 24px 32px;
        font-family: 'Plus Jakarta Sans', 'Segoe UI', Roboto, Arial, sans-serif;
        font-size: 13px;
        color: var(--ink);
        background: var(--bg);
    }

    /* Kartu judul laporan */
    .ms-head {
        display: flex; flex-wrap: wrap; align-items: center; justify-content: space-between; gap: 16px;
        margin-bottom: 16px;
        padding: 16px 20px;
        background: #fff;
        border: 1px solid var(--line);
        border-left: 4px solid var(--brand);
        border-radius: 12px;
    }
    .ms-head h1 { margin: 0 0 6px; font-size: 18px; font-weight: 700; letter-spacing: .2px; }
    .ms-meta { display: flex; flex-wrap: wrap; gap: 8px; }
    .ms-chip {
        display: inline-flex; align-items: center; gap: 6px;
        padding: 4px 12px;
        font-size: 12px; font-weight: 600;
        color: var(--brand-dark); background: var(--brand-tint);
        border-radius: 999px;
    }
    .ms-chip.month { color: var(--ink); background: var(--accent); }
    .ms-actions { display: flex; gap: 8px; }
    .ms-actions button, .ms-actions input[type="button"] {
        float: none !important;
        margin: 0 !important;
        padding: 9px 16px !important;
        font: inherit; font-size: 13px; font-weight: 600;
        color: #fff !important;
        background: var(--brand) !important;
        border: 1px solid var(--brand);
        border-radius: 8px !important;
        cursor: pointer;
    }
    .ms-actions button.secondary { color: var(--brand) !important; background: #fff !important; border-color: #b9d9d0; }
    .ms-actions button:hover, .ms-actions input[type="button"]:hover { filter: brightness(.95); }

    /* Tabel laporan */
    .ms-card {
        background: #fff;
        border: 1px solid var(--line);
        border-radius: 12px;
        overflow: auto;
        max-height: calc(100vh - 150px);
    }
    #product-table {
        border-collapse: separate;
        border-spacing: 0;
        width: 100%;
        font-size: 12.5px;
    }
    #product-table th, #product-table td {
        padding: 7px 10px;
        border: 0;
        border-bottom: 1px solid var(--line);
        text-align: right;
        vertical-align: middle;
        white-space: nowrap;
        font-variant-numeric: tabular-nums;
    }
    #product-table thead { position: sticky; top: 0; z-index: 2; }
    /* Baris judul & tanggal tetap ada di tabel (ikut Export Excel), di layar diganti kartu judul */
    #product-table tr.title-row { display: none; }
    #product-table thead th {
        font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: .4px;
        text-align: center;
        color: #fff;
        background: var(--brand);
        border-bottom: 1px solid rgba(255,255,255,.18);
        border-right: 1px solid rgba(255,255,255,.12);
        -webkit-print-color-adjust: exact; print-color-adjust: exact;
    }
    #product-table thead th.sub-stock { font-weight: 600; text-transform: none; color: #d6ebe5; background: var(--brand-dark); }
    #product-table thead th.col-fbg { background: #0f4d40; }
    #product-table td.col-no { width: 48px; text-align: center; color: var(--muted); }
    #product-table td.left { text-align: left; white-space: normal; min-width: 280px; }
    #product-table td.col-sat { text-align: center; color: var(--muted); }
    #product-table td.col-fbg { font-weight: 700; background: #f6faf9; }
    #product-table tbody tr:not(.section-row):nth-child(even) td { background-color: #fbfcfc; }
    #product-table tbody tr:not(.section-row):nth-child(even) td.col-fbg { background-color: #eff6f4; }
    #product-table tbody tr:not(.section-row):hover td { background-color: var(--brand-tint); }
    #product-table tr.section-row td {
        padding: 9px 10px;
        font-size: 12px; font-weight: 700; letter-spacing: .5px;
        text-align: left;
        color: var(--brand-dark);
        background: #eef4f2;
        border-top: 1px solid #cfdcd8;
        border-bottom: 1px solid #cfdcd8;
        -webkit-print-color-adjust: exact; print-color-adjust: exact;
    }
    #product-table tr.section-row td:first-child { text-align: center; }
    #product-table td.empty { padding: 22px 10px; text-align: center; color: var(--muted); font-style: italic; }

    @media (max-width: 640px) {
        body { padding: 12px 12px 24px; }
        .ms-head { padding: 14px 16px; }
        .ms-head h1 { font-size: 16px; }
    }
    @media print {
        @page { size: A4 landscape; margin: 10mm; }
        body { padding: 0; background: #fff; font-size: 11px; }
        .printhide, .ms-actions { display: none !important; }
        .ms-head { margin-bottom: 8px; padding: 0 0 8px; border: 0; border-bottom: 2px solid var(--brand); border-radius: 0; }
        .ms-card { max-height: none; overflow: visible; border: 0; border-radius: 0; }
        #product-table { font-size: 10px; }
        #product-table th, #product-table td { padding: 3px 6px; border: 1px solid #c9d3d0; }
        #product-table thead { position: static; display: table-header-group; }
        #product-table tr { page-break-inside: avoid; }
    }
</style>
</head>
<body>
<%
    dim datemonstock
    datemonstock = Request("datemonstock")

    bulan = Month(CDate(datemonstock))
    tahun = Year(CDate(datemonstock))
                
    If bulan = 1 Then
        namabulan = "Januari"
    ElseIf bulan = 2 Then
        namabulan = "Februari"
    ElseIf bulan = 3 Then
        namabulan = "Maret"
    ElseIf bulan = 4 Then
        namabulan = "April"
    ElseIf bulan = 5 Then
        namabulan = "Mei"
    ElseIf bulan = 6 Then
        namabulan = "Juni"
    ElseIf bulan = 7 Then
        namabulan = "Juli"
    ElseIf bulan = 8 Then
        namabulan = "Agustus"
    ElseIf bulan = 9 Then
        namabulan = "September"
    ElseIf bulan = 10 Then
        namabulan = "Oktober"
    ElseIf bulan = 11 Then
        namabulan = "November"
    ElseIf bulan = 12 Then
        namabulan = "Desember"
    Else
        namabulan = "-"
    End If

    Function SafeDbl(v)
        If IsNull(v) Or v = "" Then
            SafeDbl = 0
        Else
            SafeDbl = CDbl(v)
        End If
    End Function
    Dim cabangArray, cabangDisplay
    cabangArray = Array("PT.FBG", "PT.ARJ", "PT.BRS", "PT.GAS", "PT.KSR", "PT.BIA", "PT.BRP", "PT.KAA", "PT.CAE", "PT.ESR")
    
    Dim totalCabang()
    ReDim totalCabang(UBound(cabangArray))
    Dim j
    For j = 0 To UBound(cabangArray)
        totalCabang(j) = 0
    Next

    Dim rsPupuk, rsChemical
    sqlPupuk1 = "SELECT ROW_NUMBER() OVER (ORDER BY stockname) NO, " & _
                "stockname || ' - ' || STOCKCODE AS NAMA_BARANG, " & _
                "uom AS SAT, " & _
                "(COALESCE(ARJ_QTY, 0)+COALESCE(BRS_QTY, 0)+COALESCE(GAS_QTY, 0)+COALESCE(KSR_QTY, 0)+COALESCE(BIA_QTY, 0)+COALESCE(BRP_QTY, 0)+COALESCE(KAA_QTY, 0)+COALESCE(CAE_QTY, 0)+COALESCE(ESR_QTY, 0)) AS FBG_SHI_Stock, " & _
                "COALESCE(ARJ_QTY, 0) ARJ_SHI_Stock, " & _
                "COALESCE(BRS_QTY, 0) BRS_SHI_Stock, " & _
                "COALESCE(GAS_QTY, 0) GAS_SHI_Stock, " & _
                "COALESCE(KSR_QTY, 0) KSR_SHI_Stock, " & _
                "COALESCE(BIA_QTY, 0) BIA_SHI_Stock, " & _
                "COALESCE(BRP_QTY, 0) BRP_SHI_Stock, " & _
                "COALESCE(KAA_QTY, 0) KAA_SHI_Stock, " & _
                "COALESCE(CAE_QTY, 0) CAE_SHI_Stock, " & _
                "COALESCE(ESR_QTY, 0) ESR_SHI_Stock " & _
                "FROM " & _
                "(SELECT t.STOCKCODE, " & _
                "stock_api.get_stockname(t.STOCKCODE, t.fcba) stockname, " & _
                "stock_api.get_uom(t.STOCKCODE, t.fcba) uom, " & _
                "bu.fccompanycode, " & _
                "sum(nvl(t.openingbalanceqty, 0))+sum(nvl(t.receivenoteqty, 0))-sum(nvl(t.issueqty, 0))-sum(nvl(t.quantitytransfered, 0))+sum(nvl(t.quantityreceived, 0))+sum(nvl(t.qtywriteoff, 0)) AS total_qty " & _
                "FROM " & _
                "(SELECT DISTINCT l.fcba, " & _
                "l.STORECODE, " & _
                "l.itemcode STOCKCODE, " & _
                "l.openingbalanceqty, " & _
                "l.qtyreceived AS receivenoteqty, " & _
                "0 AS issueqty, " & _
                "0 AS quantitytransfered, " & _
                "0 AS quantityreceived, " & _
                "0 AS qtywriteoff " & _
                "FROM stockledger l " & _
                "INNER JOIN bankstock k ON (l.itemcode=k.fccode) " & _
                "WHERE k.isactive='Y' " & _
                "AND l.YEAR='"&tahun&"' " & _
                "AND l.MONTH='"&bulan&"'AND l.STORECODE = 'ST1' " & _
                "AND l.itemcode in (select distinct fccode from stock where stockgroupcode IN ('103', '305'))" & _
                "UNION ALL " & _
                "SELECT DISTINCT r.fcba, " & _
                "r.storecode, " & _
                "rd.purchaseitemcode STOCKCODE, " & _
                "0 AS openingbalanceqty, " & _
                "rd.quantity AS receivenoteqty, " & _
                "0 AS issueqty, " & _
                "0 AS quantitytransfered, " & _
                "0 AS quantityreceived, " & _
                "0 AS qtywriteoff " & _
                "FROM receivenote r, " & _
                "receivenotedetail rd " & _
                "WHERE r.receivenotecode=rd.receivenotecode AND rd.purchaseitemcode in (select distinct fccode from stock where stockgroupcode IN ('103', '305')) " & _
                "AND r.rndate BETWEEN to_date('"& datemonstock &"', 'mm/dd/yyyy') AND to_date('"& datemonstock &"', 'mm/dd/yyyy') " & _
                "AND r.POSTINGFLG = 1 " & _
                "UNION ALL " & _
                "SELECT DISTINCT h.fcba, " & _
                "h.storecode, " & _
                "d.STOCKCODE, " & _
                "0 AS openingbalanceqty, " & _
                "0 AS receivenoteqty, " & _
                "d.quantity AS issueqty, " & _
                "0 AS quantitytransfered, " & _
                "0 AS quantityreceived, " & _
                "0 AS qtywriteoff " & _
                "FROM ISSUE h, " & _
                "issuedetails d " & _
                "WHERE h.sivcode=d.sivcode AND d.STOCKCODE in (select distinct fccode from stock where stockgroupcode IN ('103', '305'))" & _
                "AND h.sivdate BETWEEN to_date('"& datemonstock &"', 'mm/dd/yyyy') AND to_date('"& datemonstock &"', 'mm/dd/yyyy') " & _
                "AND h.POSTINGFLG = 1 " & _
                "UNION ALL " & _
                "SELECT DISTINCT d.fcba, " & _
                "d.storecode, " & _
                "d.STOCKCODE, " & _
                "0 AS openingbalanceqty, " & _
                "0 AS receivenoteqty, " & _
                "0 AS issueqty, " & _
                "d.quantitytransfered, " & _
                "0 AS quantityreceived, " & _
                "0 AS qtywriteoff " & _
                "FROM transporttaskdtl_tab d, " & _
                "transporttask_tab h " & _
                "WHERE h.orderno=d.orderno " & _
                "AND d.STOCKCODE in (select distinct fccode from stock where stockgroupcode IN ('103', '305')) " & _
                "AND h.orderdate BETWEEN to_date('"& datemonstock &"', 'mm/dd/yyyy') AND to_date('"& datemonstock &"', 'mm/dd/yyyy') " & _
                "AND h.POSTINGFLG = 1 " & _
                "UNION ALL " & _
                "SELECT DISTINCT d.fcba, " & _
                "d.locationcode, " & _
                "d.STOCKCODE, " & _
                "0 AS openingbalanceqty, " & _
                "0 AS receivenoteqty, " & _
                "0 AS issueqty, " & _
                "0 AS quantitytransfered, " & _
                "d.quantityreceived AS quantityreceived, " & _
                "0 AS qtywriteoff " & _
                "FROM receivedtask_tab h, " & _
                "receivedtaskdtl_tab d " & _
                "WHERE h.orderno=d.orderno AND d.STOCKCODE in (select distinct fccode from stock where stockgroupcode IN ('103', '305')) " & _
                "AND h.orderdate BETWEEN to_date('"& datemonstock &"', 'mm/dd/yyyy') AND to_date('"& datemonstock &"', 'mm/dd/yyyy') " & _
                "AND h.POSTINGFLG = 1 " & _
                "UNION ALL " & _
                "SELECT DISTINCT h.fcba, " & _
                "h.STORECODE locationcode, " & _
                "d.stockcode, " & _
                "0 AS openingbalanceqty, " & _
                "0 AS receivenoteqty, " & _
                "0 AS issueqty, " & _
                "0 AS quantitytransfered, " & _
                "0 AS quantityreceived, " & _
                "d.quantity AS qtywriteoff " & _
                "FROM writeoff h, " & _
                "writeoffdetails d " & _
                "WHERE h.writeoffcode=d.writeoffcode AND d.STOCKCODE in (select distinct fccode from stock where stockgroupcode IN ('103', '305')) " & _
                "AND h.wodate BETWEEN to_date('"& datemonstock &"', 'mm/dd/yyyy') AND to_date('"& datemonstock &"', 'mm/dd/yyyy') " & _
                "UNION ALL " & _
                "SELECT DISTINCT h.fcba, " & _
                "h.STORE locationcode, " & _
                "itemcode stockcode, " & _
                "0 AS openingbalanceqty, " & _
                "0 AS receivenoteqty, " & _
                "0 AS issueqty, " & _
                "0 AS quantitytransfered, " & _
                "qtyreturn AS quantityreceived, " & _
                "0 AS qtywriteoff " & _
                "FROM returnissuehdr h, " & _
                "returnissuedtl d, " & _
                "stock st " & _
                "WHERE h.returnno=d.returnno AND st.stockgroupcode IN ('103', '305') " & _
                "AND h.returndate BETWEEN to_date('"& datemonstock &"', 'mm/dd/yyyy') AND to_date('"& datemonstock &"', 'mm/dd/yyyy') AND d.itemcode in (select distinct fccode from stock where stockgroupcode IN ('103', '305'))) t " & _
                "INNER JOIN businessunit bu ON (t.fcba = bu.FCCODE) " & _
                "GROUP BY t.STOCKCODE, " & _
                "stock_api.get_stockname(t.STOCKCODE, t.fcba), " & _
                "stock_api.get_uom(t.STOCKCODE, t.fcba), " & _
                "bu.fccompanycode) pivot_data " & _
                "PIVOT " & _
                "(sum(total_qty) " & _
                "FOR fccompanycode IN ('PT.ARJ' AS ARJ_QTY, " & _
                "'PT.BRS' AS BRS_QTY, " & _
                "'PT.GAS' AS GAS_QTY, " & _
                "'PT.KSR' AS KSR_QTY, " & _
                "'PT.BIA' AS BIA_QTY, " & _
                "'PT.BRP' AS BRP_QTY, " & _
                "'PT.KAA' AS KAA_QTY, " & _
                "'PT.CAE' AS CAE_QTY, " & _
                "'PT.ESR' AS ESR_QTY)) " & _
                "ORDER BY stockname"
    sqlChemical =  "SELECT ROW_NUMBER() OVER (ORDER BY stockname) NO, " & _
                "stockname || ' - ' || STOCKCODE AS NAMA_BARANG, " & _
                "uom AS SAT, " & _
                "(COALESCE(ARJ_QTY, 0)+COALESCE(BRS_QTY, 0)+COALESCE(GAS_QTY, 0)+COALESCE(KSR_QTY, 0)+COALESCE(BIA_QTY, 0)+COALESCE(BRP_QTY, 0)+COALESCE(KAA_QTY, 0)+COALESCE(CAE_QTY, 0)+COALESCE(ESR_QTY, 0)) AS FBG_SHI_Stock, " & _
                "COALESCE(ARJ_QTY, 0) ARJ_SHI_Stock, " & _
                "COALESCE(BRS_QTY, 0) BRS_SHI_Stock, " & _
                "COALESCE(GAS_QTY, 0) GAS_SHI_Stock, " & _
                "COALESCE(KSR_QTY, 0) KSR_SHI_Stock, " & _
                "COALESCE(BIA_QTY, 0) BIA_SHI_Stock, " & _
                "COALESCE(BRP_QTY, 0) BRP_SHI_Stock, " & _
                "COALESCE(KAA_QTY, 0) KAA_SHI_Stock, " & _
                "COALESCE(CAE_QTY, 0) CAE_SHI_Stock, " & _
                "COALESCE(ESR_QTY, 0) ESR_SHI_Stock " & _
                "FROM " & _
                "(SELECT t.STOCKCODE, " & _
                "stock_api.get_stockname(t.STOCKCODE, t.fcba) stockname, " & _
                "stock_api.get_uom(t.STOCKCODE, t.fcba) uom, " & _
                "bu.fccompanycode, " & _
                "sum(nvl(t.openingbalanceqty, 0))+sum(nvl(t.receivenoteqty, 0))-sum(nvl(t.issueqty, 0))-sum(nvl(t.quantitytransfered, 0))+sum(nvl(t.quantityreceived, 0))+sum(nvl(t.qtywriteoff, 0)) AS total_qty " & _
                "FROM " & _
                "(SELECT DISTINCT l.fcba, " & _
                "l.STORECODE, " & _
                "l.itemcode STOCKCODE, " & _
                "l.openingbalanceqty, " & _
                "l.qtyreceived AS receivenoteqty, " & _
                "0 AS issueqty, " & _
                "0 AS quantitytransfered, " & _
                "0 AS quantityreceived, " & _
                "0 AS qtywriteoff " & _
                "FROM stockledger l " & _
                "INNER JOIN bankstock k ON (l.itemcode=k.fccode) " & _
                "WHERE k.isactive='Y' " & _
                "AND l.YEAR='2026' " & _
                "AND l.MONTH='1'AND l.STORECODE = 'ST1' " & _
                "AND l.itemcode in (select distinct fccode from stock where stockgroupcode IN ('102', '301')) " & _
                "UNION ALL " & _
                "SELECT DISTINCT r.fcba, " & _
                "r.storecode, " & _
                "rd.purchaseitemcode STOCKCODE, " & _
                "0 AS openingbalanceqty, " & _
                "rd.quantity AS receivenoteqty, " & _
                "0 AS issueqty, " & _
                "0 AS quantitytransfered, " & _
                "0 AS quantityreceived, " & _
                "0 AS qtywriteoff " & _
                "FROM receivenote r, " & _
                "receivenotedetail rd " & _
                "WHERE r.receivenotecode=rd.receivenotecode AND rd.purchaseitemcode in (select distinct fccode from stock where stockgroupcode IN ('102', '301')) " & _
                "AND r.rndate BETWEEN to_date('"& datemonstock &"', 'mm/dd/yyyy') AND to_date('"& datemonstock &"', 'mm/dd/yyyy') " & _
                "AND r.POSTINGFLG = 1 " & _
                "UNION ALL " & _
                "SELECT DISTINCT h.fcba, " & _
                "h.storecode, " & _
                "d.STOCKCODE, " & _
                "0 AS openingbalanceqty, " & _
                "0 AS receivenoteqty, " & _
                "d.quantity AS issueqty, " & _
                "0 AS quantitytransfered, " & _
                "0 AS quantityreceived, " & _
                "0 AS qtywriteoff " & _
                "FROM ISSUE h, " & _
                "issuedetails d " & _
                "WHERE h.sivcode=d.sivcode AND d.STOCKCODE in (select distinct fccode from stock where stockgroupcode IN ('102', '301')) " & _
                "AND h.sivdate BETWEEN to_date('"& datemonstock &"', 'mm/dd/yyyy') AND to_date('"& datemonstock &"', 'mm/dd/yyyy') " & _
                "AND h.POSTINGFLG = 1 " & _
                "UNION ALL " & _
                "SELECT DISTINCT d.fcba, " & _
                "d.storecode, " & _
                "d.STOCKCODE, " & _
                "0 AS openingbalanceqty, " & _
                "0 AS receivenoteqty, " & _
                "0 AS issueqty, " & _
                "d.quantitytransfered, " & _
                "0 AS quantityreceived, " & _
                "0 AS qtywriteoff " & _
                "FROM transporttaskdtl_tab d, " & _
                "transporttask_tab h " & _
                "WHERE h.orderno=d.orderno " & _
                "AND h.rowstateid NOT IN('Rejected') AND d.STOCKCODE in (select distinct fccode from stock where stockgroupcode IN ('102', '301')) " & _
                "AND h.orderdate BETWEEN to_date('"& datemonstock &"', 'mm/dd/yyyy') AND to_date('"& datemonstock &"', 'mm/dd/yyyy') " & _
                "AND h.POSTINGFLG = 1 " & _
                "UNION ALL " & _
                "SELECT DISTINCT d.fcba, " & _
                "d.locationcode, " & _
                "d.STOCKCODE, " & _
                "0 AS openingbalanceqty, " & _
                "0 AS receivenoteqty, " & _
                "0 AS issueqty, " & _
                "0 AS quantitytransfered, " & _
                "d.quantityreceived AS quantityreceived, " & _
                "0 AS qtywriteoff " & _
                "FROM receivedtask_tab h, " & _
                "receivedtaskdtl_tab d " & _
                "WHERE h.orderno=d.orderno AND d.STOCKCODE in (select distinct fccode from stock where stockgroupcode IN ('102', '301')) " & _
                "AND h.orderdate BETWEEN to_date('"& datemonstock &"', 'mm/dd/yyyy') AND to_date('"& datemonstock &"', 'mm/dd/yyyy') " & _
                "AND h.POSTINGFLG = 1 " & _
                "UNION ALL " & _
                "SELECT DISTINCT h.fcba, " & _
                "h.STORECODE locationcode, " & _
                "d.stockcode, " & _
                "0 AS openingbalanceqty, " & _
                "0 AS receivenoteqty, " & _
                "0 AS issueqty, " & _
                "0 AS quantitytransfered, " & _
                "0 AS quantityreceived, " & _
                "d.quantity AS qtywriteoff " & _
                "FROM writeoff h, " & _
                "writeoffdetails d " & _
                "WHERE h.writeoffcode=d.writeoffcode AND d.STOCKCODE in (select distinct fccode from stock where stockgroupcode IN ('102', '301')) " & _
                "AND h.wodate BETWEEN to_date('"& datemonstock &"', 'mm/dd/yyyy') AND to_date('"& datemonstock &"', 'mm/dd/yyyy') " & _
                "UNION ALL " & _
                "SELECT DISTINCT h.fcba, " & _
                "h.STORE locationcode, " & _
                "itemcode stockcode, " & _
                "0 AS openingbalanceqty, " & _
                "0 AS receivenoteqty, " & _
                "0 AS issueqty, " & _
                "0 AS quantitytransfered, " & _
                "qtyreturn AS quantityreceived, " & _
                "0 AS qtywriteoff " & _
                "FROM returnissuehdr h, " & _
                "returnissuedtl d, " & _
                "stock st " & _
                "WHERE h.returnno=d.returnno " & _
                "AND h.returndate BETWEEN to_date('"& datemonstock &"', 'mm/dd/yyyy') AND to_date('"& datemonstock &"', 'mm/dd/yyyy') AND d.itemcode in (select distinct fccode from stock where stockgroupcode IN ('102', '301'))) t " & _
                "INNER JOIN businessunit bu ON (t.fcba = bu.FCCODE) " & _
                "GROUP BY t.STOCKCODE, " & _
                "stock_api.get_stockname(t.STOCKCODE, t.fcba), " & _
                "stock_api.get_uom(t.STOCKCODE, t.fcba), " & _
                "bu.fccompanycode) pivot_data " & _
                "PIVOT " & _
                "(sum(total_qty) " & _
                "FOR fccompanycode IN ('PT.ARJ' AS ARJ_QTY, " & _
                "'PT.BRS' AS BRS_QTY, " & _
                "'PT.GAS' AS GAS_QTY, " & _
                "'PT.KSR' AS KSR_QTY, " & _
                "'PT.BIA' AS BIA_QTY, " & _
                "'PT.BRP' AS BRP_QTY, " & _
                "'PT.KAA' AS KAA_QTY, " & _
                "'PT.CAE' AS CAE_QTY, " & _
                "'PT.ESR' AS ESR_QTY)) " & _
                "ORDER BY stockname"
    ' Response.write sqlPupuk1
    Set rsPupuk = Server.CreateObject("ADODB.Recordset")
    rsPupuk.Open sqlPupuk1, connstr
	
    Set rsChemical = Server.CreateObject("ADODB.Recordset")
    rsChemical.Open sqlChemical, connstr

    Function DisplaySatuan(satuan)
        If satuan = "KG" Then
            DisplaySatuan = "TON"
        Else
            DisplaySatuan = satuan
        End If
    End Function

    Function FormatStockValue(nilai, satuan)
        Dim convertedValue
        Dim result
        convertedValue = SafeDbl(nilai)
        
        If satuan = "KG" Then
            ' Konversi KG ke TON (1 TON = 1000 KG)
            convertedValue = convertedValue / 1000
            result = Round(convertedValue, 2)
            FormatStockValue = result
        Else
            ' Untuk satuan lainnya, tampilkan apa adanya
            FormatStockValue = convertedValue
        End If
    End Function
%>
<div class="ms-head">
    <div>
        <h1>Monitoring Stock Pupuk &amp; Agrochemical <%=tahun%></h1>
        <div class="ms-meta">
            <span class="ms-chip">Tanggal: <%=datemonstock%></span>
            <span class="ms-chip month"><%=UCase(namabulan)%>&nbsp;<%=tahun%></span>
            <span class="ms-chip">Satuan KG ditampilkan dalam TON</span>
        </div>
    </div>
    <div class="ms-actions printhide">
        <button type="button" class="secondary" onclick="window.print()">Print</button>
        <!--#include file="print_excel.asp"-->
    </div>
</div>
<div class="ms-card">
<table id="product-table" class="xthf-site">
    <thead>
    <!-- Judul -->
    <tr class="title-row">
        <th colspan="13" class="title">
            MONITORING STOCK PUPUK & AGROCHEMICAL <%=tahun%>
        </th>
    </tr>

    <!-- Tanggal -->
    <tr class="title-row">
        <th class="date">
            <%=datemonstock%>
        </th>
        <th><%=UCase(namabulan)%></th>
        <th colspan="11"></th>
    </tr>

    <!-- Header utama -->
    <tr>
        <th rowspan="2" class="header-blue">NO</th>
        <th rowspan="2" class="header-blue">NAMA BARANG</th>
        <th rowspan="2" class="header-blue">SAT</th>

        <th class="header-lokasi col-fbg">FBG<br>SHI</th>
        <th class="header-lokasi">ARJ<br>SHI</th>
        <th class="header-lokasi">BRS<br>SHI</th>
        <th class="header-lokasi">GAS<br>SHI</th>
        <th class="header-lokasi">KSR<br>SHI</th>
        <th class="header-lokasi">BIA<br>SHI</th>
        <th class="header-lokasi">BRP<br>SHI</th>
        <th class="header-lokasi">KAA<br>SHI</th>
        <th class="header-lokasi">CAE<br>SHI</th>
        <th class="header-lokasi">ESR<br>SHI</th>
    </tr>

    <!-- Sub header -->
    <tr>
        <th class="sub-stock col-fbg">Stock</th>
        <th class="sub-stock">Stock</th>
        <th class="sub-stock">Stock</th>
        <th class="sub-stock">Stock</th>
        <th class="sub-stock">Stock</th>
        <th class="sub-stock">Stock</th>
        <th class="sub-stock">Stock</th>
        <th class="sub-stock">Stock</th>
        <th class="sub-stock">Stock</th>
        <th class="sub-stock">Stock</th>
    </tr>
    </thead>
    <tbody>

    <!-- Section Pupuk -->
    <tr class="section-row">
        <td class="section">A.</td>
        <td colspan="12" class="section">PUPUK</td>
    </tr>
    <%
        if Not rsPupuk.eof then
            nCountofField=rsPupuk.fields.count-1
            no=1
            While Not rsPupuk.EOF
    %>
        <tr>
            <td class="col-no"><%=no%></td>
            <td class="left"><%=rsPupuk("NAMA_BARANG")%></td>
            <td class="col-sat"><%=DisplaySatuan(rsPupuk("SAT"))%></td>
            <td class="col-fbg"><%=FormatStockValue(rsPupuk("FBG_SHI_Stock"), rsPupuk("SAT"))%></td>
            <td><%=FormatStockValue(rsPupuk("ARJ_SHI_Stock"), rsPupuk("SAT"))%></td>
            <td><%=FormatStockValue(rsPupuk("BRS_SHI_Stock"), rsPupuk("SAT"))%></td>
            <td><%=FormatStockValue(rsPupuk("GAS_SHI_Stock"), rsPupuk("SAT"))%></td>
            <td><%=FormatStockValue(rsPupuk("KSR_SHI_Stock"), rsPupuk("SAT"))%></td>
            <td><%=FormatStockValue(rsPupuk("BIA_SHI_Stock"), rsPupuk("SAT"))%></td>
            <td><%=FormatStockValue(rsPupuk("BRP_SHI_Stock"), rsPupuk("SAT"))%></td>
            <td><%=FormatStockValue(rsPupuk("KAA_SHI_Stock"), rsPupuk("SAT"))%></td>
            <td><%=FormatStockValue(rsPupuk("CAE_SHI_Stock"), rsPupuk("SAT"))%></td>
            <td><%=FormatStockValue(rsPupuk("ESR_SHI_Stock"), rsPupuk("SAT"))%></td>
        </tr>
    <%
        rsPupuk.MoveNext
        no=no+1
    Wend
    else
    response.Write("<tr><td colspan=""13"" class=""empty"">Data is empty</td></tr>")
    end if
    rsPupuk.Close
    set rsPupuk = Nothing
    %>

    <!-- Section Agrochemical -->
    <tr class="section-row">
        <td class="section">B.</td>
        <td colspan="12" class="section">AGROCHEMICAL</td>
    </tr>

    <%
        if Not rsChemical.eof then
            nCountofField=rsChemical.fields.count-1
            no=1
            While Not rsChemical.EOF
    %>
        <tr>
            <td class="col-no"><%=no%></td>
            <td class="left"><%=rsChemical("NAMA_BARANG")%></td>
            <td class="col-sat"><%=DisplaySatuan(rsChemical("SAT"))%></td>
            <td class="col-fbg"><%=FormatStockValue(rsChemical("FBG_SHI_Stock"), rsChemical("SAT"))%></td>
            <td><%=FormatStockValue(rsChemical("ARJ_SHI_Stock"), rsChemical("SAT"))%></td>
            <td><%=FormatStockValue(rsChemical("BRS_SHI_Stock"), rsChemical("SAT"))%></td>
            <td><%=FormatStockValue(rsChemical("GAS_SHI_Stock"), rsChemical("SAT"))%></td>
            <td><%=FormatStockValue(rsChemical("KSR_SHI_Stock"), rsChemical("SAT"))%></td>
            <td><%=FormatStockValue(rsChemical("BIA_SHI_Stock"), rsChemical("SAT"))%></td>
            <td><%=FormatStockValue(rsChemical("BRP_SHI_Stock"), rsChemical("SAT"))%></td>
            <td><%=FormatStockValue(rsChemical("KAA_SHI_Stock"), rsChemical("SAT"))%></td>
            <td><%=FormatStockValue(rsChemical("CAE_SHI_Stock"), rsChemical("SAT"))%></td>
            <td><%=FormatStockValue(rsChemical("ESR_SHI_Stock"), rsChemical("SAT"))%></td>
        </tr>
    <%
        rsChemical.MoveNext
        no=no+1
    Wend
    else
    response.Write("<tr><td colspan=""13"" class=""empty"">Data is empty</td></tr>")
    end if
    rsChemical.Close
    set rsChemical = Nothing
    %>

    </tbody>
</table>
</div>

</body>
</html>
