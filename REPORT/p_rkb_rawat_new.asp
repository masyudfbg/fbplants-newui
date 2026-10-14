<% ValidateSession %>
<!--#include virtual="/ConndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->

<%
Dim v_txtMonth, v_txtYear
Dim rsDefaultBa, strDefaultBa, varDefaultBa
Dim rscurrentdate, strcurrentdate, varDefaultYear, varDefaultMonth
Dim ScriptName, v_txtsectionname, v_txtGangCode, v_txtStatus
Dim rsBaInfo, strBaInfo, varFcbaName, varCompanyName, varManagerName, varKTUName, no
DIM totalbiaya

v_txtYear    = Request("tahun")
v_txtMonth   = Request("bulan")
v_txtRegion  = Request("txtRegion")
v_txtcompany = Request("txtCompany")
v_txtfcba    = Request("fcba")

If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
    v_txtFcba = varDefaultBa
End If

strBaInfo = " SELECT" & _
            "     FCNAME as FCBANAME" & _
            "     , FCCOMPANYNAME as COMPANYNAME" & _
            "     , FCMANAGER" & _
            "     , FCKTU" & _
            " FROM BUSINESSUNIT BU" & _
            " WHERE" & _
            "     FCCODE = '" & v_txtFcba & "'"

Set rsBaInfo = Server.CreateObject("ADODB.Recordset")
rsBaInfo.Open strBaInfo, connstr, adOpenStatic

If Not rsBaInfo.EOF Then
    varFcbaName   = rsBaInfo("FCBANAME")
    varCompanyName = rsBaInfo("COMPANYNAME")
    varManagerName = rsBaInfo("FCMANAGER")
    varKTUName    = rsBaInfo("FCKTU")
Else
    varFcbaName   = ""
    varCompanyName = ""
    varManagerName = ""
    varKTUName    = ""
End If
rsBaInfo.Close
Set rsBaInfo = Nothing

' ============================================================
' PASS 1 : Baca semua data ke array & hitung rowspan
' ============================================================
Dim sqldtl, rsdtl

sqldtl = "select " & _
    " h.FCBA, " & _
    " to_char(f.PLANTINGDATE,'yyyy') tahuntanam, " & _
    " h.rkbno, " & _
    " u.actcode, " & _
    " j.fcname, " & _
    " u.locationcode, " & _
    " NVL(f.LUAS_GROSS,0) LUAS_GROSS, " & _
    " NVL(f.TOTALSTAND,0) TOTALSTAND, " & _
    " CASE WHEN Nvl(rkb_rowno,0)=0 OR (rkb_rowno=1 AND bhn_count>=1) THEN " & _
    "   NVL(u.VOLUME,0) " & _
    " ELSE 0 END AS VOLUME, " & _
    " j.UNITOFMEASUREMENT, " & _
    " NVL(u.NORMAMANDAYS,0) NORMAMANDAYS, " & _
    " CASE WHEN Nvl(rkb_rowno,0)=0 OR (rkb_rowno=1 AND bhn_count>=1) THEN " & _
    "   NVL(u.MANDAYS,0) " & _
    " ELSE 0 END AS MANDAYS, " & _
    " CASE WHEN Nvl(rkb_rowno,0)=0 OR (rkb_rowno=1 AND bhn_count>=1) THEN " & _
    "   NVL(u.TOTALAMOUNT_HK,0) " & _
    " ELSE 0 END AS TOTALAMOUNT_HK, " & _
    " CASE WHEN Nvl(rkb_rowno,0)=0 OR (rkb_rowno=1 AND bhn_count>=1) THEN " & _
    "   NVL(u.RATEBORONGAN,0) " & _
    " ELSE 0 END AS RATEBORONGAN, " & _
    " CASE WHEN Nvl(rkb_rowno,0)=0 OR (rkb_rowno=1 AND bhn_count>=1) THEN " & _
    "   NVL(u.TOTALAMOUNT_BORONGAN,0) " & _
    " ELSE 0 END AS TOTALAMOUNT_BORONGAN, " & _
    " A.FCNAME NAMABAHAN, " & _
    " b.UOM, " & _
    " b.ACTIVEMATERIALNORM, " & _
    " NVL(b.qty,0) QTY, " & _
    " NVL(b.RATESTOCK,0) RATESTOCK, " & _
    " NVL(b.TOTALAMOUNT,0) TOTALAMOUNTBAHAN, " & _
    " CASE WHEN Nvl(rkb_rowno,0)=0 OR (rkb_rowno=1 AND bhn_count>=1) THEN " & _
    "   NVL(u.TOTALAMOUNT_HK,0) + NVL(TOTALAMOUNT_BORONGAN,0) + NVL(summarybahan,0) " & _
    " ELSE 0 END AS AMOUNTHKBAHAN, " & _
    " b.rkb_rowno, " & _
    " b.bhn_count " & _
    " from rkbhdr h " & _
    " inner join rkbupahdtl u on h.rkbno=u.rkbno and h.fcba=u.fcba " & _
    " left join ( " & _
    "   SELECT b.*, " & _
    "   ROW_NUMBER() OVER (PARTITION BY rkbno,actcode,locationcode ORDER BY rkbno,actcode,locationcode,STOCKCODE ASC) rkb_rowno, " & _
    "   Count(rkbno) OVER (PARTITION BY rkbno,actcode,locationcode ORDER BY rkbno,actcode,locationcode ASC) bhn_count " & _
    "   FROM RKBBAHANDTL b " & _
    " ) b on (u.rkbno=b.rkbno and u.fcba=b.fcba AND u.actcode=b.actcode AND u.locationcode=b.locationcode) " & _
    " left join ( " & _
    "   select rkbno,fcba,actcode,locationcode, sum(totalamount) as summarybahan " & _
    "   from RKBBAHANDTL group by rkbno,fcba,actcode,locationcode " & _
    " ) ba on ba.RKBNO=h.rkbno AND ba.fcba=u.fcba AND u.actcode=ba.actcode AND u.locationcode=ba.locationcode " & _
    " inner join field_hist_tab f on u.locationcode=f.FCCODE and u.FCBA=f.FCBA " & _
    " inner join jobmaster j on u.actcode=j.fccode " & _
    " LEFT JOIN ACTIVEMATERIAL a ON b.STOCKCODE=A.FCCODE " & _
    " where h.DOCUMENT_CLASSIFICATION='501' " & _
    "   AND h.fcba='" & v_txtfcba & "' " & _
    "   AND f.YEAR='" & v_txtYear & "' " & _
    "   and f.MONTH='" & v_txtMonth & "' " & _
    "   AND h.PERIOD_YEAR='" & v_txtYear & "' " & _
    "   and h.PERIOD_MONTH='" & v_txtMonth & "' " & _
    " ORDER BY to_char(f.PLANTINGDATE,'yyyy'),h.rkbno,u.actcode,u.locationcode,b.STOCKCODE"

Set rsdtl = connStr.Execute(sqldtl)

' Tampung semua baris ke array
' Kolom: 0=FCBA,1=tahuntanam,2=rkbno,3=actcode,4=fcname,5=locationcode,
'        6=LUAS_GROSS,7=TOTALSTAND,8=VOLUME,9=UNITOFMEASUREMENT,
'        10=NORMAMANDAYS,11=MANDAYS,12=TOTALAMOUNT_HK,13=RATEBORONGAN,
'        14=TOTALAMOUNT_BORONGAN,15=NAMABAHAN,16=UOM,17=ACTIVEMATERIALNORM,
'        18=QTY,19=RATESTOCK,20=TOTALAMOUNTBAHAN,21=AMOUNTHKBAHAN

Dim arrData()
Dim totalRows
totalRows = 0

If Not rsdtl.EOF Then
    rsdtl.MoveFirst
    Do While Not rsdtl.EOF
        ReDim Preserve arrData(21, totalRows)
        arrData(0,  totalRows) = rsdtl("FCBA")
        arrData(1,  totalRows) = rsdtl("tahuntanam")
        arrData(2,  totalRows) = rsdtl("rkbno")
        arrData(3,  totalRows) = rsdtl("actcode")
        arrData(4,  totalRows) = rsdtl("fcname")
        arrData(5,  totalRows) = rsdtl("locationcode")
        arrData(6,  totalRows) = rsdtl("LUAS_GROSS")
        arrData(7,  totalRows) = rsdtl("TOTALSTAND")
        arrData(8,  totalRows) = rsdtl("VOLUME")
        arrData(9,  totalRows) = rsdtl("UNITOFMEASUREMENT")
        arrData(10, totalRows) = rsdtl("NORMAMANDAYS")
        arrData(11, totalRows) = rsdtl("MANDAYS")
        arrData(12, totalRows) = rsdtl("TOTALAMOUNT_HK")
        arrData(13, totalRows) = rsdtl("RATEBORONGAN")
        arrData(14, totalRows) = rsdtl("TOTALAMOUNT_BORONGAN")
        arrData(15, totalRows) = rsdtl("NAMABAHAN")
        arrData(16, totalRows) = rsdtl("UOM")
        arrData(17, totalRows) = rsdtl("ACTIVEMATERIALNORM")
        arrData(18, totalRows) = rsdtl("QTY")
        arrData(19, totalRows) = rsdtl("RATESTOCK")
        arrData(20, totalRows) = rsdtl("TOTALAMOUNTBAHAN")
        arrData(21, totalRows) = rsdtl("AMOUNTHKBAHAN")
        totalRows = totalRows + 1
        rsdtl.MoveNext
    Loop
End If
rsdtl.Close
Set rsdtl = Nothing

' -------------------------------------------------------
' Hitung rowspan untuk masing-masing kolom merge:
'   rsKebun   -> merge berdasarkan: FCBA + tahuntanam + rkbno + actcode  (kolom Kebun/Pabrik)
'   rsTanam   -> merge berdasarkan: FCBA + tahuntanam + rkbno + actcode  (kolom Tahun Tanam)
'   rsRkbno   -> merge berdasarkan: FCBA + tahuntanam + rkbno + actcode  (kolom Nomor RKB)
'   rsActcode -> merge berdasarkan: FCBA + tahuntanam + rkbno + actcode  (kolom Kode Pekerjaan)
'
' Semua 4 kolom tersebut di-merge dengan kunci yang sama:
'   FCBA + tahuntanam + rkbno + actcode
' -------------------------------------------------------
Dim rsKebun(), rsTanam(), rsRkbno(), rsActcode()

If totalRows > 0 Then
    ReDim rsKebun(totalRows - 1)
    ReDim rsTanam(totalRows - 1)
    ReDim rsRkbno(totalRows - 1)
    ReDim rsActcode(totalRows - 1)

    Dim i, j
    For i = 0 To totalRows - 1
        ' Kunci merge = FCBA + tahuntanam + rkbno + actcode
        Dim keyI
        keyI = arrData(0,i) & "||" & arrData(1,i) & "||" & arrData(2,i) & "||" & arrData(3,i)

        ' Hitung berapa baris berurutan ke depan yang punya kunci sama
        Dim cnt
        cnt = 0
        For j = i To totalRows - 1
            Dim keyJ
            keyJ = arrData(0,j) & "||" & arrData(1,j) & "||" & arrData(2,j) & "||" & arrData(3,j)
            If keyJ = keyI Then
                cnt = cnt + 1
            Else
                Exit For
            End If
        Next

        ' Cek apakah baris ini adalah baris pertama kunci ini (baris sebelumnya beda kunci)
        Dim isFirst
        isFirst = True
        If i > 0 Then
            Dim keyPrev
            keyPrev = arrData(0,i-1) & "||" & arrData(1,i-1) & "||" & arrData(2,i-1) & "||" & arrData(3,i-1)
            If keyPrev = keyI Then isFirst = False
        End If

        If isFirst Then
            rsKebun(i)   = cnt
            rsTanam(i)   = cnt
            rsRkbno(i)   = cnt
            rsActcode(i) = cnt
        Else
            rsKebun(i)   = 0   ' 0 = jangan cetak td (sudah di-merge)
            rsTanam(i)   = 0
            rsRkbno(i)   = 0
            rsActcode(i) = 0
        End If
    Next
End If
%>
<html>
<head>
    <title>FBPlants Reporting - Rencana Kerja Bulanan Rawat</title>
    <link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
    <link rel="stylesheet" href="/css/jquery-ui.css" />
    <script src="/js/jquery-1.9.1.js"></script>
    <script src="/js/jquery-ui.js"></script>
    <script language="JavaScript" src='/js/x.js'></script>
    <script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
    <script language="JavaScript">
    xAddEventListener(window, 'load', function() {
        new xTableHeaderFixed('xthf-site', window);
    }, false);
    </script>
    <script language="JavaScript" src="/js/date-picker.js"></script>

<style type="text/css">
@media all {
#footer_note {width:65%;margin:auto;margin-top:20px;}
#footer_note .block {font-size:10px;float:left;width:20%;height:90px;background-color:#fff;margin-left:1%;text-align:center;}
input[type="text"].isian {border:0px;box-shadow:none;padding:0px;font-size:10px;color:#000;height:20px;width:100%;text-align:center;font-weight:bold;}
hr{border-top:1px #000000 solid;border-bottom:0px;border-left:0px;border-right:0px;}

/* Merge cell style */
td.merge-cell {
    vertical-align: middle;
    text-align: center;
    font-weight: bold;
}

/* Subtotal row */
.subtotal-row td {
    background-color: #F0B429 !important;
    font-weight: bold;
}
}
@media print {
#footer_note {margin-top:20px;}
#footer_note .block {font-size:10px;float:left;width:20%;height:90px;background-color:#fff;margin-left:1%;text-align:center;}
input[type="text"].isian {border:0px;box-shadow:none;padding:0px;font-size:10px;color:#000;height:20px;width:100%;text-align:center;font-weight:bold;}
hr{border-top:1px #000000 solid;border-bottom:0px;border-left:0px;border-right:0px;}
.printhide {display:none;}
td.merge-cell {vertical-align:middle;text-align:center;font-weight:bold;}
.subtotal-row td {background-color:#F0B429 !important;font-weight:bold;}
@page { size:Letter landscape; }
}
</style>
<link href="/plugins/chosen/chosen.min.css" rel="stylesheet" />
</head>

<body>
<div id="wrapper">
    <div id="header-wrapper">
        <div class="printhide">
            <input type="button" value="Print" onClick="window.print()" style="font-family:'FontAwesome';margin-top:-10px;">
            <!--#include file="../print_excel.asp"-->
        </div>
    </div>

    <table border="0" width="100%" id="product-table" class="xthf-site">
        <thead>
            <tr>
                <th colspan="30" style="background-color:#FFFFFF;border:none;">
                    <div id="header-wrapper">
                        <div id="head_judul" style="margin-bottom:-30px;position:static;text-align:left;">
                            <h4>PT. <%=varCompanyName%></h4>
                            <h5><%=varFcbaName%></h5>
                        </div>
                        <div id="subheader" class="centered">
                            <h2>LAPORAN RENCANA KERJA BULANAN RAWAT</h2>
                            <br/>
                            <h4>PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
                        </div>
                    </div>
                </th>
            </tr>
            <tr>
                <th rowspan="3" style="color:#fff;">NO</th>
                <th rowspan="3" style="color:#fff;">Kebun/Pabrik</th>
                <th rowspan="3" style="color:#fff;">Tahun Tanam</th>
                <th rowspan="3" style="color:#fff;">Nomor RKB</th>
                <th rowspan="3" style="color:#fff;">Kode Pekerjaan</th>
                <th rowspan="3" style="color:#fff;">Nama Pekerjaan</th>
                <th colspan="5" style="color:#fff;">Lokasi</th>
                <th colspan="5" style="color:#fff;">Tenaga Kerja</th>
                <th colspan="6" style="color:#fff;">Pemakaian Bahan/ Alokasi Unit</th>
                <th rowspan="3" style="color:#fff;">Total Biaya (HK + Bahan)</th>
                <th rowspan="3" style="color:#fff;">Cost/Satuan</th>
            </tr>
            <tr>
                <th rowspan="2" style="color:#fff;">Blok</th>
                <th rowspan="2" style="color:#fff;">Luas</th>
                <th rowspan="2" style="color:#fff;">Pokok</th>
                <th rowspan="2" style="color:#fff;">Luas Dikerjakan</th>
                <th rowspan="2" style="color:#fff;">Satuan</th>
                <th colspan="3" style="color:#fff;">Harian</th>
                <th colspan="2" style="color:#fff;">Borongan / Premi</th>
                <th rowspan="2" style="color:#fff;">Nama Bahan / Unit</th>
                <th rowspan="2" style="color:#fff;">Satuan Bahan</th>
                <th rowspan="2" style="color:#fff;">Norma Bahan/HM</th>
                <th rowspan="2" style="color:#fff;">Jumlah Bahan</th>
                <th rowspan="2" style="color:#fff;">Harga Satuan(Rp)</th>
                <th rowspan="2" style="color:#fff;">Biaya Bahan(Rp)</th>
            </tr>
            <tr>
                <th style="color:#fff;">Norma HK/HA/KG</th>
                <th style="color:#fff;">Jumlah HK</th>
                <th style="color:#fff;">Jumlah Biaya (Rp)</th>
                <th style="color:#fff;">Harga Satuan</th>
                <th style="color:#fff;">Jumlah Biaya Borongan/Premi (Rp)</th>
            </tr>
        </thead>
        <tbody>
<%
' ============================================================
' PASS 2 : Cetak HTML menggunakan data dari array + rowspan
' ============================================================
Dim costsatuan
Dim prev_rkbno, prev_actcode
Dim sub_Luas, sub_Pokok, sub_LuasDikerjakan, sub_Mandays, sub_Totalamount, sub_Boronganpremi
Dim sub_Qty, sub_Totalamountstock, sub_Totalbiaya, sub_Costsatuan
Dim sub_NamaKerja

Dim gt_LuasDikerjakan, gt_Mandays, gt_Totalamount, gt_Boronganpremi
Dim gt_Qty, gt_Totalamountstock, gt_Totalbiaya, gt_Costsatuan

prev_rkbno           = ""
prev_actcode         = ""
sub_Luas 	 		 = 0
sub_Pokok 	 		 = 0
sub_LuasDikerjakan   = 0
sub_Mandays          = 0
sub_Totalamount      = 0
sub_Boronganpremi    = 0
sub_Qty              = 0
sub_Totalamountstock = 0
sub_Totalbiaya       = 0
sub_Costsatuan       = 0
sub_NamaKerja        = ""

gt_LuasDikerjakan    = 0
gt_Mandays           = 0
gt_Totalamount       = 0
gt_Boronganpremi     = 0
gt_Qty               = 0
gt_Totalamountstock  = 0
gt_Totalbiaya        = 0
gt_Costsatuan        = 0

no = 1

Dim r
For r = 0 To totalRows - 1

    Dim cur_fcba, cur_tanam, cur_rkbno, cur_actcode, cur_fcname
    cur_fcba    = arrData(0, r)
    cur_tanam   = arrData(1, r)
    cur_rkbno   = arrData(2, r)
    cur_actcode = arrData(3, r)
    cur_fcname  = arrData(4, r)

    ' --- Deteksi pergantian grup & cetak subtotal ---
    If (prev_rkbno <> "" Or prev_actcode <> "") And _
       (cur_rkbno <> prev_rkbno Or cur_actcode <> prev_actcode) Then
%>
            <tr class="subtotal-row">
                <td align="center" colspan="7"><strong>SUBTOTAL</strong></td>
				<td align="right"><strong><%=formatnumber(sub_Luas,2)%></strong></td>
                <td align="right"><strong><%=formatnumber(sub_Pokok,2)%></strong></td>
                <td align="right"><strong><%=formatnumber(sub_LuasDikerjakan,2)%></strong></td>
                <td align="right" colspan="2"></td>
                <td align="right"><strong><%=formatnumber(sub_Mandays,2)%></strong></td>
                <td align="right"><strong><%=formatnumber(sub_Totalamount,2)%></strong></td>
                <td align="right"></td>
                <td align="right"><strong><%=formatnumber(sub_Boronganpremi,2)%></strong></td>
                <td align="right" colspan="3"></td>
                <td align="center"><strong><%=formatnumber(sub_Qty,2)%></strong></td>
                <td align="right"></td>
                <td align="right"><strong><%=formatnumber(sub_Totalamountstock,2)%></strong></td>
                <td align="right"><strong><%=formatnumber(sub_Totalbiaya,2)%></strong></td>
                <td align="right"><strong><%=formatnumber(sub_Costsatuan,2)%></strong></td>
            </tr>
<%
        sub_Luas			 = 0
        sub_Pokok			 = 0
        sub_LuasDikerjakan   = 0
        sub_Mandays          = 0
        sub_Totalamount      = 0
        sub_Boronganpremi    = 0
        sub_Qty              = 0
        sub_Totalamountstock = 0
        sub_Totalbiaya       = 0
        sub_Costsatuan       = 0
    End If

    prev_rkbno   = cur_rkbno
    prev_actcode = cur_actcode
    sub_NamaKerja = cur_fcname

    ' Hitung costsatuan
    If CDbl(to_nol(arrData(8, r))) = 0 Then
        costsatuan = 0
    Else
        costsatuan = CDbl(to_nol(arrData(21, r))) / CDbl(to_nol(arrData(8, r)))
    End If
%>
            <tr>
                <td align="center"><%=no%></td>

                <%' --- Kebun/Pabrik: merge berdasarkan kunci grup ---
                If rsKebun(r) > 0 Then%>
                <td align="center" class="merge-cell" rowspan="<%=rsKebun(r)%>"><%=cur_fcba%></td>
                <%End If%>

                <%' --- Tahun Tanam ---
                If rsTanam(r) > 0 Then%>
                <td align="center" class="merge-cell" rowspan="<%=rsTanam(r)%>"><%=cur_tanam%></td>
                <%End If%>

                <%' --- Nomor RKB ---
                If rsRkbno(r) > 0 Then%>
                <td align="left" class="merge-cell" rowspan="<%=rsRkbno(r)%>"><%=cur_rkbno%></td>
                <%End If%>

                <%' --- Kode Pekerjaan ---
                If rsActcode(r) > 0 Then%>
                <td align="left" class="merge-cell" rowspan="<%=rsActcode(r)%>"><%=cur_actcode%></td>
                <%End If%>

                <td align="left"><%=cur_fcname%></td>
                <td align="left"><%=arrData(5, r)%></td>
                <td align="right"><%=formatnumber(arrData(6, r),2)%></td>
                <td align="right"><%=formatnumber(arrData(7, r),2)%></td>
                <td align="right"><%=formatnumber(arrData(8, r),2)%></td>
                <td align="left"><%=arrData(9, r)%></td>
                <td align="center"><%=formatnumber(arrData(10, r),2)%></td>
                <td align="right"><%=formatnumber(arrData(11, r),2)%></td>
                <td align="right"><%=formatnumber(arrData(12, r),2)%></td>
                <td align="right"><%=formatnumber(arrData(13, r),2)%></td>
                <td align="right"><%=formatnumber(arrData(14, r),2)%></td>
                <td align="left"><%=arrData(15, r)%></td>
                <td align="left"><%=arrData(16, r)%></td>
                <td align="center"><%=arrData(17, r)%></td>
                <td align="center"><%=formatnumber(arrData(18, r),2)%></td>
                <td align="right"><%=formatnumber(arrData(19, r),2)%></td>
                <td align="right"><%=formatnumber(arrData(20, r),2)%></td>
                <td align="right"><%=formatnumber(arrData(21, r),2)%></td>
                <td align="right"><%=formatnumber(costsatuan,2)%></td>
            </tr>
<%
    ' Akumulasi subtotal
    sub_Luas   			 = CDbl(sub_Luas)   + CDbl(arrData(6,  r))
	sub_Pokok  			 = CDbl(sub_Pokok)   + CDbl(arrData(7,  r))
    sub_LuasDikerjakan   = CDbl(sub_LuasDikerjakan)   + CDbl(arrData(8,  r))
    sub_Mandays          = CDbl(sub_Mandays)           + CDbl(arrData(11, r))
    sub_Totalamount      = CDbl(sub_Totalamount)       + CDbl(arrData(12, r))
    sub_Boronganpremi    = CDbl(sub_Boronganpremi)     + CDbl(arrData(14, r))
    sub_Qty              = CDbl(sub_Qty)               + CDbl(arrData(18, r))
    sub_Totalamountstock = CDbl(sub_Totalamountstock)  + CDbl(arrData(20, r))
    sub_Totalbiaya       = CDbl(sub_Totalbiaya)        + CDbl(arrData(21, r))
    sub_Costsatuan       = CDbl(sub_Costsatuan)        + CDbl(costsatuan)

    ' Akumulasi grand total
    gt_LuasDikerjakan    = CDbl(gt_LuasDikerjakan)    + CDbl(arrData(8,  r))
    gt_Mandays           = CDbl(gt_Mandays)            + CDbl(arrData(11, r))
    gt_Totalamount       = CDbl(gt_Totalamount)        + CDbl(arrData(12, r))
    gt_Boronganpremi     = CDbl(gt_Boronganpremi)      + CDbl(arrData(14, r))
    gt_Qty               = CDbl(gt_Qty)               + CDbl(arrData(18, r))
    gt_Totalamountstock  = CDbl(gt_Totalamountstock)   + CDbl(arrData(20, r))
    gt_Totalbiaya        = CDbl(gt_Totalbiaya)         + CDbl(arrData(21, r))
    gt_Costsatuan        = CDbl(gt_Costsatuan)         + CDbl(costsatuan)

    no = no + 1
Next

' Subtotal untuk grup terakhir
If prev_rkbno <> "" Then
%>
            <tr class="subtotal-row">
                <td align="center" colspan="7"><strong>SUBTOTAL</strong></td>
                <td align="right"><strong><%=formatnumber(sub_Pokok,2)%></strong></td>
                <td align="right"><strong><%=formatnumber(sub_Luas,2)%></strong></td>
				<td align="right"><strong><%=formatnumber(sub_LuasDikerjakan,2)%></strong></td>
                <td align="right"></td>
                <td align="right"></td>
                <td align="right"><strong><%=formatnumber(sub_Mandays,2)%></strong></td>
                <td align="right"><strong><%=formatnumber(sub_Totalamount,2)%></strong></td>
                <td align="right"></td>
                <td align="right"><strong><%=formatnumber(sub_Boronganpremi,2)%></strong></td>
                <td align="right" colspan="3"></td>
                <td align="center"><strong><%=formatnumber(sub_Qty,2)%></strong></td>
                <td align="right"></td>
                <td align="right"><strong><%=formatnumber(sub_Totalamountstock,2)%></strong></td>
                <td align="right"><strong><%=formatnumber(sub_Totalbiaya,2)%></strong></td>
                <td align="right"><strong><%=formatnumber(sub_Costsatuan,2)%></strong></td>
            </tr>
<%End If%>

            <tr style="background-color:#27AE60;">
                <td align="center" colspan="7"><strong>Grand Total</strong></td>
				<td align="right"><strong><%=formatnumber(sub_Pokok,2)%></strong></td>
				<td align="right"><strong><%=formatnumber(sub_Luas,2)%></strong></td>
                <td align="right"><strong><%=formatnumber(gt_LuasDikerjakan,2)%></strong></td>
                <td align="right"></td>
                <td align="right"></td>
                <td align="right"><strong><%=formatnumber(gt_Mandays,2)%></strong></td>
                <td align="right"><strong><%=formatnumber(gt_Totalamount,2)%></strong></td>
                <td align="right"></td>
                <td align="right"><strong><%=formatnumber(gt_Boronganpremi,2)%></strong></td>
                <td align="right" colspan="3"></td>
                <td align="center"><strong><%=formatnumber(gt_Qty,2)%></strong></td>
                <td align="right"></td>
                <td align="right"><strong><%=formatnumber(gt_Totalamountstock,2)%></strong></td>
                <td align="right"><strong><%=formatnumber(gt_Totalbiaya,2)%></strong></td>
                <td align="right"><strong><%=formatnumber(gt_Costsatuan,2)%></strong></td>
            </tr>
        </tbody>
    </table>
</div>

<script src="/plugins/chosen/chosen.jquery.min.js"></script>
<script type="text/javascript">
    var config = {
        '.chosen-select'           : {},
        '.chosen-select-deselect'  : {allow_single_deselect:true},
        '.chosen-select-no-single' : {disable_search_threshold:10},
        '.chosen-select-no-results': {no_results_text:'Oops, nothing found!'},
        '.chosen-select-width'     : {width:"95%"}
    }
    for (var selector in config) {
        $(selector).chosen(config[selector]);
    }
</script>
</body>
</html>
