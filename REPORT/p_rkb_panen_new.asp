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

v_txtYear  = Request("tahun")
v_txtMonth = Request("bulan")
v_txtfcba  = Request("fcba")

strDefaultBa = " select" & _
               "     fccode as fcba" & _
               " from businessunit bu" & _
               " inner join user_fcba_tab uft on (bu.fccode = uft.fcba) " & _
               " where isdefault = 'TRUE'" & _
               "       and fcusercode = '" & Session("CodeName") & "'"

Set rsDefaultBa = Server.CreateObject("ADODB.Recordset")
rsDefaultBa.Open strDefaultBa, connstr, adOpenStatic

If Not rsDefaultBa.EOF Then
    varDefaultBa = rsDefaultBa("fcba")
Else
    varDefaultBa = Session("fcba")
End If
rsDefaultBa.Close
Set rsDefaultBa = Nothing

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
    varFcbaName    = rsBaInfo("FCBANAME")
    varCompanyName = rsBaInfo("COMPANYNAME")
    varManagerName = rsBaInfo("FCMANAGER")
    varKTUName     = rsBaInfo("FCKTU")
Else
    varFcbaName    = ""
    varCompanyName = ""
    varManagerName = ""
    varKTUName     = ""
End If
rsBaInfo.Close
Set rsBaInfo = Nothing

txtComp = Request("txtCompany")

' ============================================================
' PASS 1 : Baca semua data ke array
' Kolom array:
'   0=FCBA, 1=tahuntanam, 2=rkbno, 3=actcode, 4=fcname,
'   5=locationcode, 6=LUAS_GROSS, 7=TOTALSTAND,
'   8=UNITOFMEASUREMENT, 9=ABW_BUDGET, 10=ROTASI,
'   11=kg_budget, 12=tonasekg, 13=NORMA_HK (dari ACTIVITY_NORM_HARV),
'   14=RATEHK, 15=NORMABASIS_JJG, 16=RATE_SB, 17=RATE_LB
' ============================================================
Dim sqldtl1, rsdtl1

sqldtl1 = "select h.FCBA,u.tahuntanam,h.rkbno,u.actcode,j.fcname,u.locationcode," & _
          " NVL(f.LUAS_GROSS,0)LUAS_GROSS,NVL(f.TOTALSTAND,0)TOTALSTAND," & _
          " j.UNITOFMEASUREMENT,nvl(u.ABW_BUDGET,0)ABW_BUDGET,NVL(u.ROTASI,0)ROTASI," & _
          " nvl(u.kg_budget,0)kg_budget,(nvl(u.ABW_BUDGET,0)*NVL(u.kg_budget,0))tonasekg," & _
          " nvl(u.NORMAMANDAYS,0)NORMAMANDAYS,nvl(u.MANDAYS,0)MANDAYS," & _
          " nvl(u.TOTALAMOUNT_HK,0)TOTALAMOUNT_HK,nvl(u.RATE,0)RATEHK," & _
          " NVL(u.BASIS_JJG,0)BASIS_JJG,NVL(u.RATE_SB,0)RATE_SB,NVL(u.RATE_LB,0)RATE_LB," & _
          " a.ROTASI ROTASI,NVL(a.NORMA_HK,0)NORMA_HK,NVL(a.BASIS_JJG,0)NORMABASIS_JJG" & _
          " from rkbhdr h" & _
          " inner join rkbupahdtl u on h.rkbno=u.rkbno and h.fcba=u.fcba" & _
          " inner join field_hist_tab f on u.locationcode=f.FCCODE and u.FCBA=f.FCBA" & _
          " inner join jobmaster j on u.actcode=j.fccode" & _
          " inner join ACTIVITY_NORM_HARV a on u.tahuntanam=a.PLANTINGYEAR AND u.fcba=a.fcba" & _
          " where h.DOCUMENT_CLASSIFICATION='503'" & _
          "   AND h.fcba='" & v_txtfcba & "'" & _
          "   AND f.YEAR='" & v_txtYear & "'" & _
          "   and f.MONTH='" & v_txtMonth & "'" & _
          "   AND h.PERIOD_YEAR='" & v_txtYear & "'" & _
          "   and h.PERIOD_MONTH='" & v_txtMonth & "'" & _
          "   AND to_char(a.TGLBERLAKU,'yyyy')>2023" & _
          " order by u.tahuntanam, h.rkbno"

Set rsdtl1 = connStr.Execute(sqldtl1)

Dim arrData()
Dim totalRows
totalRows = 0

If Not rsdtl1.EOF Then
    Do While Not rsdtl1.EOF
        ReDim Preserve arrData(17, totalRows)
        arrData(0,  totalRows) = rsdtl1("FCBA")
        arrData(1,  totalRows) = rsdtl1("tahuntanam")
        arrData(2,  totalRows) = rsdtl1("rkbno")
        arrData(3,  totalRows) = rsdtl1("actcode")
        arrData(4,  totalRows) = rsdtl1("fcname")
        arrData(5,  totalRows) = rsdtl1("locationcode")
        arrData(6,  totalRows) = rsdtl1("LUAS_GROSS")
        arrData(7,  totalRows) = rsdtl1("TOTALSTAND")
        arrData(8,  totalRows) = rsdtl1("UNITOFMEASUREMENT")
        arrData(9,  totalRows) = rsdtl1("ABW_BUDGET")
        arrData(10, totalRows) = rsdtl1("ROTASI")
        arrData(11, totalRows) = rsdtl1("kg_budget")
        arrData(12, totalRows) = rsdtl1("tonasekg")
        arrData(13, totalRows) = rsdtl1("NORMA_HK")
        arrData(14, totalRows) = rsdtl1("RATEHK")
        arrData(15, totalRows) = rsdtl1("NORMABASIS_JJG")
        arrData(16, totalRows) = rsdtl1("RATE_SB")
        arrData(17, totalRows) = rsdtl1("RATE_LB")
        totalRows = totalRows + 1
        rsdtl1.MoveNext
    Loop
End If
rsdtl1.Close
Set rsdtl1 = Nothing

' ============================================================
' Hitung rowspan untuk kolom merge:
'   Kunci merge = FCBA + tahuntanam + rkbno + actcode
' ============================================================
Dim rsKebun(), rsTanam(), rsRkbno(), rsActcode()

If totalRows > 0 Then
    ReDim rsKebun(totalRows - 1)
    ReDim rsTanam(totalRows - 1)
    ReDim rsRkbno(totalRows - 1)
    ReDim rsActcode(totalRows - 1)

    Dim i, j
    For i = 0 To totalRows - 1
        Dim keyI
        keyI = arrData(0,i) & "||" & arrData(1,i) & "||" & arrData(2,i) & "||" & arrData(3,i)

        ' Hitung baris berurutan dengan kunci sama
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

        ' Cek apakah baris pertama dari kunci ini
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
            rsKebun(i)   = 0  ' 0 = jangan cetak td (sudah di-merge)
            rsTanam(i)   = 0
            rsRkbno(i)   = 0
            rsActcode(i) = 0
        End If
    Next
End If
%>
<html>
<head>
    <title>FBPlants Reporting - Rencana Kerja Bulanan Panen</title>
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

/* Merge cell */
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
                            <h2>LAPORAN RENCANA KERJA BULANAN PANEN</h2>
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
                <th rowspan="3" style="color:#fff;">Nomor Dokumen</th>
                <th rowspan="3" style="color:#fff;">Kode Pekerjaan</th>
                <th rowspan="3" style="color:#fff;">Nama Pekerjaan</th>
                <th colspan="3" style="color:#fff;">Blok</th>
                <th rowspan="3" style="color:#fff;">Satuan</th>
                <th rowspan="3" style="color:#fff;">BJR</th>
                <th rowspan="3" style="color:#fff;">Rotasi</th>
                <th colspan="2" style="color:#fff;">Pencapaian Hasil</th>
                <th colspan="2" style="color:#fff;">Tenaga Kerja</th>
                <th rowspan="3" style="color:#fff;">Total HK (RP)</th>
                <th colspan="4" style="color:#fff;">Panen</th>
                <th colspan="2" style="color:#fff;">Total Biaya Panen</th>
            </tr>
            <tr>
                <th rowspan="2" style="color:#fff;">Blok</th>
                <th rowspan="2" style="color:#fff;">Luas(Ha)</th>
                <th rowspan="2" style="color:#fff;">Jumlah Pokok</th>
                <th rowspan="2" style="color:#fff;">JJG</th>
                <th rowspan="2" style="color:#fff;">KG</th>
                <th rowspan="2" style="color:#fff;">Norma HK</th>
                <th rowspan="2" style="color:#fff;">HK</th>
                <th rowspan="2" style="color:#fff;">JJG Basis</th>
                <th rowspan="2" style="color:#fff;">Sampai Basis (SB) (Rp)</th>
                <th rowspan="2" style="color:#fff;">Premi Lebih Basis (LB) (Rp)</th>
                <th rowspan="2" style="color:#fff;">Total Premi</th>
                <th rowspan="2" style="color:#fff;">Total</th>
                <th rowspan="2" style="color:#fff;">RP/KG</th>
            </tr>
            <tr></tr>
        </thead>
        <tbody>
<%
' ============================================================
' PASS 2 : Cetak HTML dari array + rowspan
' ============================================================
Dim totalpremibasis, hk, amounthkrp, totalsemua, rpperkg

' Subtotal per grup
Dim prev_rkbno, prev_actcode, sub_NamaKerja
Dim sub_Luasblok, sub_jmlpokok, sub_jjg, sub_kg, sub_hk, sub_totalhk
Dim sub_jjgbasis, sub_sampaibasis, sub_premilbbasis, sub_totalpremi, sub_total, sub_rpkg

' Grand total
Dim gt_Luasblok, gt_jmlpokok, gt_jjg, gt_kg, gt_hk, gt_totalhk
Dim gt_jjgbasis, gt_sampaibasis, gt_premilbbasis, gt_totalpremi, gt_total, gt_rpkg

prev_rkbno   = ""
prev_actcode = ""
sub_NamaKerja      = ""
sub_Luasblok       = 0
sub_jmlpokok       = 0
sub_jjg            = 0
sub_kg             = 0
sub_hk             = 0
sub_totalhk        = 0
sub_jjgbasis       = 0
sub_sampaibasis    = 0
sub_premilbbasis   = 0
sub_totalpremi     = 0
sub_total          = 0
sub_rpkg           = 0

gt_Luasblok        = 0
gt_jmlpokok        = 0
gt_jjg             = 0
gt_kg              = 0
gt_hk              = 0
gt_totalhk         = 0
gt_jjgbasis        = 0
gt_sampaibasis     = 0
gt_premilbbasis    = 0
gt_totalpremi      = 0
gt_total           = 0
gt_rpkg            = 0

no = 1

Dim r
For r = 0 To totalRows - 1

    Dim cur_fcba, cur_tanam, cur_rkbno, cur_actcode, cur_fcname
    cur_fcba    = arrData(0, r)
    cur_tanam   = arrData(1, r)
    cur_rkbno   = arrData(2, r)
    cur_actcode = arrData(3, r)
    cur_fcname  = arrData(4, r)

    ' Hitung nilai kalkulasi baris ini
    totalpremibasis = CDbl(arrData(16, r)) + CDbl(arrData(17, r))
    hk              = CDbl(arrData(6, r))  * CDbl(arrData(10, r)) * CDbl(arrData(13, r))
    amounthkrp      = CDbl(hk)             * CDbl(arrData(14, r))
    totalsemua      = CDbl(amounthkrp)     + CDbl(totalpremibasis)
    If CDbl(arrData(12, r)) = 0 Then
        rpperkg = 0
    Else
        rpperkg = CDbl(totalsemua) / CDbl(arrData(12, r))
    End If

    ' --- Deteksi pergantian grup -> cetak subtotal ---
    If (prev_rkbno <> "" Or prev_actcode <> "") And _
       (cur_rkbno <> prev_rkbno Or cur_actcode <> prev_actcode) Then
%>
            <tr class="subtotal-row">
                <td align="center" colspan="7"><strong>SUBTOTAL</strong></td>
                <td align="right"><strong><%=formatnumber(sub_Luasblok,2)%></strong></td>
                <td align="right"><strong><%=formatnumber(sub_jmlpokok,2)%></strong></td>
                <td align="right"></td>
                <td align="right"></td>
                <td align="right"></td>
                <td align="right"><strong><%=formatnumber(sub_jjg,2)%></strong></td>
                <td align="right"><strong><%=formatnumber(sub_kg,2)%></strong></td>
                <td align="right"></td>
                <td align="right"><strong><%=formatnumber(sub_hk,2)%></strong></td>
                <td align="right"><strong><%=formatnumber(sub_totalhk,2)%></strong></td>
                <td align="right"><strong><%=formatnumber(sub_jjgbasis,2)%></strong></td>
                <td align="right"><strong><%=formatnumber(sub_sampaibasis,2)%></strong></td>
                <td align="right"><strong><%=formatnumber(sub_premilbbasis,2)%></strong></td>
                <td align="right"><strong><%=formatnumber(sub_totalpremi,2)%></strong></td>
                <td align="right"><strong><%=formatnumber(sub_total,2)%></strong></td>
                <td align="right"><strong><%=formatnumber(sub_rpkg,2)%></strong></td>
            </tr>
<%
        ' Reset subtotal
        sub_Luasblok     = 0
        sub_jmlpokok     = 0
        sub_jjg          = 0
        sub_kg           = 0
        sub_hk           = 0
        sub_totalhk      = 0
        sub_jjgbasis     = 0
        sub_sampaibasis  = 0
        sub_premilbbasis = 0
        sub_totalpremi   = 0
        sub_total        = 0
        sub_rpkg         = 0
    End If

    prev_rkbno   = cur_rkbno
    prev_actcode = cur_actcode
    sub_NamaKerja = cur_fcname
%>
            <tr>
                <td align="center"><%=no%></td>

                <%' --- Kebun/Pabrik ---
                If rsKebun(r) > 0 Then%>
                <td align="center" class="merge-cell" rowspan="<%=rsKebun(r)%>"><%=cur_fcba%></td>
                <%End If%>

                <%' --- Tahun Tanam ---
                If rsTanam(r) > 0 Then%>
                <td align="center" class="merge-cell" rowspan="<%=rsTanam(r)%>"><%=cur_tanam%></td>
                <%End If%>

                <%' --- Nomor Dokumen ---
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
                <td align="right"><%=formatnumber(arrData(7, r),0)%></td>
                <td align="left"><%=arrData(8, r)%></td>
                <td align="right"><%=formatnumber(arrData(9, r),2)%></td>
                <td align="right"><%=arrData(10, r)%></td>
                <td align="right"><%=formatnumber(arrData(11, r),0)%></td>
                <td align="right"><%=formatnumber(arrData(12, r),0)%></td>
                <td align="right"><%=formatnumber(arrData(13, r),2)%></td>
                <td align="right"><%=formatnumber(hk,0)%></td>
                <td align="right"><%=formatnumber(amounthkrp,0)%></td>
                <td align="right"><%=formatnumber(arrData(15, r),0)%></td>
                <td align="right"><%=formatnumber(arrData(16, r),0)%></td>
                <td align="right"><%=formatnumber(arrData(17, r),0)%></td>
                <td align="right"><%=formatnumber(totalpremibasis,0)%></td>
                <td align="right"><%=formatnumber(totalsemua,0)%></td>
                <td align="right"><%=formatnumber(rpperkg,0)%></td>
            </tr>
<%
    ' Akumulasi subtotal
    sub_Luasblok     = CDbl(sub_Luasblok)     + CDbl(arrData(6, r))
    sub_jmlpokok     = CDbl(sub_jmlpokok)     + CDbl(arrData(7, r))
    sub_jjg          = CDbl(sub_jjg)          + CDbl(arrData(11, r))
    sub_kg           = CDbl(sub_kg)           + CDbl(arrData(12, r))
    sub_hk           = CDbl(sub_hk)           + CDbl(hk)
    sub_totalhk      = CDbl(sub_totalhk)      + CDbl(amounthkrp)
    sub_jjgbasis     = CDbl(sub_jjgbasis)     + CDbl(arrData(15, r))
    sub_sampaibasis  = CDbl(sub_sampaibasis)  + CDbl(arrData(16, r))
    sub_premilbbasis = CDbl(sub_premilbbasis) + CDbl(arrData(17, r))
    sub_totalpremi   = CDbl(sub_totalpremi)   + CDbl(totalpremibasis)
    sub_total        = CDbl(sub_total)        + CDbl(totalsemua)
    sub_rpkg         = CDbl(sub_rpkg)         + CDbl(rpperkg)

    ' Akumulasi grand total
    gt_Luasblok      = CDbl(gt_Luasblok)      + CDbl(arrData(6, r))
    gt_jmlpokok      = CDbl(gt_jmlpokok)      + CDbl(arrData(7, r))
    gt_jjg           = CDbl(gt_jjg)           + CDbl(arrData(11, r))
    gt_kg            = CDbl(gt_kg)            + CDbl(arrData(12, r))
    gt_hk            = CDbl(gt_hk)            + CDbl(hk)
    gt_totalhk       = CDbl(gt_totalhk)       + CDbl(amounthkrp)
    gt_jjgbasis      = CDbl(gt_jjgbasis)      + CDbl(arrData(15, r))
    gt_sampaibasis   = CDbl(gt_sampaibasis)   + CDbl(arrData(16, r))
    gt_premilbbasis  = CDbl(gt_premilbbasis)  + CDbl(arrData(17, r))
    gt_totalpremi    = CDbl(gt_totalpremi)    + CDbl(totalpremibasis)
    gt_total         = CDbl(gt_total)         + CDbl(totalsemua)
    gt_rpkg          = CDbl(gt_rpkg)          + CDbl(rpperkg)

    no = no + 1
Next

' Subtotal grup terakhir
If prev_rkbno <> "" Then
%>
            <tr class="subtotal-row">
                <td align="center" colspan="7"><strong>SUBTOTAL</strong></td>
                <td align="right"><strong><%=formatnumber(sub_Luasblok,2)%></strong></td>
                <td align="right"><strong><%=formatnumber(sub_jmlpokok,2)%></strong></td>
                <td align="right"></td>
                <td align="right"></td>
                <td align="right"></td>
                <td align="right"><strong><%=formatnumber(sub_jjg,2)%></strong></td>
                <td align="right"><strong><%=formatnumber(sub_kg,2)%></strong></td>
                <td align="right"></td>
                <td align="right"><strong><%=formatnumber(sub_hk,2)%></strong></td>
                <td align="right"><strong><%=formatnumber(sub_totalhk,2)%></strong></td>
                <td align="right"><strong><%=formatnumber(sub_jjgbasis,2)%></strong></td>
                <td align="right"><strong><%=formatnumber(sub_sampaibasis,2)%></strong></td>
                <td align="right"><strong><%=formatnumber(sub_premilbbasis,2)%></strong></td>
                <td align="right"><strong><%=formatnumber(sub_totalpremi,2)%></strong></td>
                <td align="right"><strong><%=formatnumber(sub_total,2)%></strong></td>
                <td align="right"><strong><%=formatnumber(sub_rpkg,2)%></strong></td>
            </tr>
<%End If%>

            <tr style="background-color:#27AE60;">
                <td align="center" colspan="7"><strong>Grand Total</strong></td>
                <td align="right"><strong><%=formatnumber(gt_Luasblok,2)%></strong></td>
                <td align="right"><strong><%=formatnumber(gt_jmlpokok,2)%></strong></td>
                <td align="right"></td>
                <td align="right"></td>
                <td align="right"></td>
                <td align="right"><strong><%=formatnumber(gt_jjg,2)%></strong></td>
                <td align="right"><strong><%=formatnumber(gt_kg,2)%></strong></td>
                <td align="right"></td>
                <td align="right"><strong><%=formatnumber(gt_hk,2)%></strong></td>
                <td align="right"><strong><%=formatnumber(gt_totalhk,2)%></strong></td>
                <td align="right"><strong><%=formatnumber(gt_jjgbasis,2)%></strong></td>
                <td align="right"><strong><%=formatnumber(gt_sampaibasis,2)%></strong></td>
                <td align="right"><strong><%=formatnumber(gt_premilbbasis,2)%></strong></td>
                <td align="right"><strong><%=formatnumber(gt_totalpremi,2)%></strong></td>
                <td align="right"><strong><%=formatnumber(gt_total,2)%></strong></td>
                <td align="right"><strong><%=formatnumber(gt_rpkg,2)%></strong></td>
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
