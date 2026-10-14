<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/ConndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
'=======================================================
' BLOCK 1 — Handle POST: simpan PRICE_TENGKULAK / PRICE_TBS
'=======================================================
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then

  Dim pFcba, pDocketId, pField, pValue

  pFcba     = Trim(Request.Form("fcba"))
  pDocketId = Trim(Request.Form("docketid"))
  pField    = Trim(Request.Form("field"))
  pValue    = Trim(Request.Form("value"))

  If pFcba = "" Or pDocketId = "" Or pField = "" Then
    Response.Status = "400 Bad Request"
    Response.Write "Parameter fcba, docketid, dan field wajib diisi."
    Response.End
  End If

  ' Whitelist field — cegah SQL injection via nama field
  If pField <> "PRICE_TENGKULAK" And pField <> "PRICE_TBS" And pField <> "RATEPPN" Then
    Response.Status = "400 Bad Request"
    Response.Write "Field tidak diizinkan: " & Server.HtmlEncode(pField)
    Response.End
  End If

  ' Validasi nilai: harus angka, tidak negatif
  If Not IsNumeric(pValue) Then
    Response.Status = "400 Bad Request"
    Response.Write "Nilai harus berupa angka."
    Response.End
  End If

  Dim dblValue
  dblValue = CDbl(pValue)
  If dblValue < 0 Then
    Response.Status = "400 Bad Request"
    Response.Write "Nilai tidak boleh negatif."
    Response.End
  End If

  ' Sanitasi FCBA dan DOCKETID — tolak karakter berbahaya
  Dim badChars, bi
  badChars = Array("'", """", ";", "--", "/*", "*/", "xp_", "=", "<", ">")
  For bi = 0 To UBound(badChars)
    If InStr(pFcba,     badChars(bi)) > 0 Or _
       InStr(pDocketId, badChars(bi)) > 0 Then
      Response.Status = "400 Bad Request"
      Response.Write "Karakter tidak valid pada parameter."
      Response.End
    End If
  Next

  ' Jalankan UPDATE — pakai objek connStr yang sudah terbuka dari ConndB1.inc
  Dim sqlUpdate

  On Error Resume Next

  sqlUpdate = "UPDATE LHPTIKETHDR" & _
              " SET " & pField & " = " & dblValue & _
              " WHERE FCBA = '" & pFcba & "'" & _
              " AND DOCKETID = '" & pDocketId & "'"

  connStr.Execute sqlUpdate

  If Err.Number <> 0 Then
    Dim errSave : errSave = Err.Description
    On Error GoTo 0
    Response.Status = "500 Internal Server Error"
    Response.Write "Gagal UPDATE: " & errSave
    Response.End
  End If

  On Error GoTo 0

  Response.ContentType = "text/plain"
  Response.Write "OK"
  Response.End

End If
'=======================================================
' BLOCK 2 — GET: render halaman monitoring
'=======================================================
%>
<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Monitoring Produksi TBS Per Mobil Per Hari – Januari 2026</title>
<style>
* {
	margin: 0;
	padding: 0;
	box-sizing: border-box
}

body {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 13px;
	background: #c8c8c8;
	color: #000
}

.page {
	background: #fff;
	padding: 8px 10px 24px;
	margin: 0 auto;
	min-width: max-content
}

.co {
	font-size: 13px;
	font-weight: bold;
	line-height: 1.7;
	margin-bottom: 6px
}

.titlebar {
	display: flex;
	align-items: stretch;
	border: 1px solid #555;
	width: fit-content;
	margin-bottom: 2px
}

.t-bia {
	background: #499e10;
	border-right: 1px solid #555;
	padding: 3px 10px;
	font-weight: bold;
  color: white;
	font-size: 13px;
	text-align: center;
	line-height: 1.5;
	display: flex;
	align-items: center;
	white-space: pre-line
}

.t-main {
	background: #499e10;
	padding: 4px 16px;
	font-weight: bold;
	font-size: 13px;
  color: white;
	display: flex;
	align-items: center;
	letter-spacing: .2px
}

.wrap {
	overflow-x: auto;
	overflow-y: auto;
	max-height: calc(100vh - 160px);
	margin-top: 0;
	position: relative
}

table {
	border-collapse: collapse;
	font-size: 13px;
	table-layout: fixed
}

th,
td {
	border: 1px solid #9e9e9e;
	padding: 2px 3px;
	text-align: center;
	vertical-align: middle;
	overflow: hidden
}

/* ── STICKY HEADER ROW 1 ── */
thead {
	position: sticky;
	top: 0;
	z-index: 3;
	background: #499e10;
  color: white;
	position: sticky;
	top: 0;
	z-index: 3;
	background: #499e10;
  color: white;
  padding: 10px;
}

thead tr:nth-child(1) th.gh-pk {
	position: sticky;
	top: 0;
	z-index: 3;
	background: #499e10;
  color: white;
}

thead tr:nth-child(1) th.gh-dp {
	position: sticky;
	top: 0;
	z-index: 3;
	background: #499e10;
  color: white;
}

thead tr:nth-child(1) th.gh-cm {
	position: sticky;
	top: 0;
	z-index: 3;
	background: #499e10;
  color: white;
}

/* ── STICKY HEADER ROW 2 ── */
thead tr:nth-child(2) th.sh-kb {
	position: sticky;
	top: 29px;
	z-index: 3;
	background: #499e10;
  color: white;
}

thead tr:nth-child(2) th.sh-tk {
	position: sticky;
	top: 29px;
	z-index: 3;
	background: #499e10;
  color: white;
}

thead tr:nth-child(2) th.sh-pk {
	position: sticky;
	top: 29px;
	z-index: 3;
	background: #499e10;
  color: white;
}

thead tr:nth-child(2) th.sh-dp {
	position: sticky;
	top: 29px;
	z-index: 3;
	background: #499e10;
  color: white;
}

thead tr:nth-child(2) th.sh-cm {
	position: sticky;
	top: 29px;
	z-index: 3;
	background: #499e10;
  color: white;
}

/* ── STICKY HEADER ROW 3 ── */
thead tr:nth-child(3) th.lr {
	position: sticky;
	top: 87px;
	z-index: 3;
}

thead tr:nth-child(3) th.lr-cy {
	position: sticky;
	top: 87px;
	z-index: 3;
	background: #00B0F0
}

/* ── STICKY KOLOM PERTAMA (No.) ── */
tbody tr td:first-child {
	position: sticky;
	left: 0;
	z-index: 2;
	background: #f2f2f2
}

tfoot tr td:first-child {
	position: sticky;
	left: 0;
	z-index: 2;
	background: #FFC000
}

thead tr:nth-child(1) th:first-child {
	position: sticky;
	top: 0;
	left: 0;
	z-index: 6;
	background: #499e10
}

thead tr:nth-child(2) th:first-child {
	position: sticky;
	top: 29px;
	left: 0;
	z-index: 6;
	background: #499e10
}

thead tr:nth-child(3) th:first-child {
	position: sticky;
	top: 87px;
	left: 0;
	z-index: 6;
}

.gh-no,
.gh-kb,
.gh-pks {
	background: #499e10;
	font-weight: bold
}

.gh-tk {
	background: #499e10;
	font-weight: bold
}

.gh-pk {
	background: #499e10;
	font-weight: bold
}

.gh-dp {
	background: #499e10;
	font-weight: bold
}

.gh-cm {
	background: #499e10;
	font-weight: bold
}

.sh-kb {
	background: #499e10
}

.sh-tk {
	background: #499e10
}

.sh-pk {
	background: #499e10
}

.sh-dp {
	background: #499e10
}

.sh-cm {
	background: #499e10;
}

.lr {
	background: #fff;
	font-weight: bold;
	color: #555
}

.lr-cy {
	background: #00B0F0;
	font-weight: bold
}

tbody tr:nth-child(odd) td {
	background: #fff
}

tbody tr:nth-child(even) td {
	background: #f9f9f9
}

tbody tr:hover td {
	background: #FFFDE7 !important
}

.td-no {
	background: #f2f2f2 !important;
	color: #888;
	width: 24px
}

.td-kb {
	font-weight: bold;
	text-align: left;
	padding: 5px
}

.cde { color: #0070C0 }
.bia { color: #7030A0 }
.ghe { color: #375623 }

.td-r {
	text-align: right;
	padding-right: 4px
}

.td-l {
	text-align: left;
	padding-left: 3px
}

.td-cy {
	text-align: right;
	padding-right: 4px;
	font-weight: bold
}

tr.tot td {
	background: #FFC000 !important;
	font-weight: bold;
	border-color: #fff
}

tr.tot .td-no {
	background: #FFC000 !important;
	color: #000
}

/* ── INPUT HARGA ── */
input.inp-rp {
  width: 100%;
  height: 30px;
  border: 1px solid #ccc;
  border-radius: 2px;
  padding: 1px 3px;
  font-size: 12px;
  font-family: Arial, sans-serif;
  text-align: right;
  background: #fffde7;
  box-sizing: border-box;
}
input.inp-rp:focus {
  outline: 2px solid #1D6F42;
  background: #fff;
}

/* ── State warna input: saving / saved / error ── */
input.inp-rp.state-saving {
  background: #fff9c4;
  border-color: #f9a825;
}
input.inp-rp.state-saved {
  background: #e8f5e9;
  border-color: #388e3c;
  transition: background 1.5s ease, border-color 1.5s ease;
}
input.inp-rp.state-error {
  background: #ffebee;
  border-color: #c62828;
}

/* ── Auto-fill highlight ── */
input.inp-rp.filled-auto {
  background: #d4edda;
  transition: background 1.2s ease;
}

/* ── Status bar ── */
#statusBar {
  position: fixed;
  bottom: 12px;
  right: 16px;
  z-index: 9999;
  display: flex;
  flex-direction: column;
  gap: 4px;
  pointer-events: none;
}
.status-msg {
  padding: 6px 14px;
  border-radius: 4px;
  font-size: 12px;
  font-family: Arial, sans-serif;
  font-weight: bold;
  box-shadow: 0 2px 6px rgba(0,0,0,.25);
  opacity: 1;
  transition: opacity .6s ease;
}
.status-msg.ok  { background: #e8f5e9; color: #2e7d32; border: 1px solid #a5d6a7; }
.status-msg.err { background: #ffebee; color: #b71c1c; border: 1px solid #ef9a9a; }
.status-msg.inf { background: #e3f2fd; color: #1565c0; border: 1px solid #90caf9; }
</style>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
</head>
<body style="background-color:#fff">

<!-- ── Status bar notifikasi ── -->
<div id="statusBar"></div>

<div style="margin:20px">
  <button onclick="downloadExcel()" style="
    background:#1D6F42;color:#fff;border:none;padding:6px 16px;
    font-size:13px;font-family:Arial,sans-serif;cursor:pointer;
    border-radius:3px;display:inline-flex;align-items:center;gap:6px">
    Download Excel
  </button>
</div>

<div class="page">
  <%
    Dim sqlCompany, setCompany
    sqlCompany = "SELECT * FROM BUSINESSUNIT WHERE fccode = '" & Request("fcba") & "'"
    Set setCompany = Server.CreateObject("ADODB.Recordset")
    setCompany.Open sqlCompany, connStr, adOpenStatic
  %>
  <div class="co">FIRST BORNEO GROUP<br>PT. <%=setCompany("FCCOMPANYNAME")%></div>
  <div class="titlebar">
    <div class="t-bia"><%=setCompany("FCCOMPANYCODE")%>&#10;<%=Year(CDate(Request("datemontbsto")))%></div>
    <div class="t-main">
      MONITORING PRODUKSI TBS PER MOBIL PER HARI<br>
      Periode : <%=Request("datemontbsfrom")%> s/d <%=Request("datemontbsto")%>
    </div>
  </div>

  <%
    Dim sql, setdata, tglfrom, tglto
    tglfrom = Request("datemontbsfrom")
    tglto   = Request("datemontbsto")

    sql = " SELECT ROW_NUMBER() OVER (ORDER BY res.TGLTIMBANG ASC) AS NO, " & _
          "     res.*, " & _
          "     (res.NETTOTENGKULAK - res.BRUTO) AS TBSSUSUTKG, " & _
          "     ROUND(((res.NETTOTENGKULAK - res.BRUTO) / NULLIF(res.BRUTO, 0)), 2) AS TBSSUSUTPER " & _
          " FROM ( " & _
          "     SELECT a.FCBA, " & _
          " a.DOCKETID, " & _
          " NVL(MAX(b.TANDANQTY),0) AS TANDANQTY, " & _
          " 0 AS ESTIMASITONASE, " & _
          " NVL(MAX(c.PKSNAME), '-') AS PKSTUJUAN, " & _
          " TO_CHAR(a.DOCKETDATE, 'DD/MM/YYYY') AS TGLTIMBANG, " & _
          " a.SPBNO, " & _
          " MIN(a.TIMEIN) AS JAMMASUK, " & _
          " MAX(a.TIMEOUT) AS JAMKELUAR, " & _
          " a.NOREG AS NOPOL, " & _
          " a.DRIVERNAME AS NAMASUPIR, " & _
          " case when NVL(MAX(a.WEIGHTIN),0)  = '' THEN 0 ELSE NVL(MAX(a.WEIGHTIN),0) END  AS TARRA, " & _
          " case when NVL(MAX(a.WEIGHTOUT),0) = '' THEN 0 ELSE NVL(MAX(a.WEIGHTOUT),0) END  AS BRUTO, " & _
          " MAX(a.WEIGHTOUT) - MAX(a.WEIGHTIN) AS NETTO, " & _
          " MAX(f.WBT_CHITNO) AS NOTIKET, " & _
          " TO_CHAR(MAX(f.tglin), 'DD/MM/YYYY') AS TGLMASUKPKS, " & _
          " TO_CHAR(MAX(f.tglout), 'DD/MM/YYYY') AS TGLOUTPKS, " & _
          " case when NVL(MAX(f.WEIGHTIN),0) = '' then 0 else NVL(MAX(f.WEIGHTIN),0) end AS BERATINPKS, " & _
          " case when NVL(MAX(f.WEIGHTOUT),0) = '' then 0 else NVL(MAX(f.WEIGHTOUT),0) end AS BERATOUTPKS, " & _
          " NVL(MAX(f.WEIGHTIN),0) - NVL(MAX(f.WEIGHTOUT),0) AS NETTOPKS, " & _
          " 0 AS POTPKSPER, " & _
          " MAX(F.POTONGAN) AS POTPKSKG, " & _
          " '0' AS NETTOPKS2, " & _
          " a.NOREG AS NOPOLPKS, " & _
          " a.DRIVERNAME AS DRIVERPKS, " & _
          " 0 - MAX(a.WEIGHTIN) AS TBSPULANG, " & _
          " NVL(MAX(f.ACCEPTQUANTITY), 0) AS NETTOTENGKULAK, " & _
          " 0 AS SELISIH, " & _
          " '0' AS TBSBALIKPER, " & _
          " (CASE  " & _
          " WHEN REGEXP_LIKE(TRIM(MAX(f.QUANTITY)), '^\d+(\.\d+)?$') THEN TO_NUMBER(MAX(f.QUANTITY))  " & _
          " ELSE 0  " & _
          " END) -  " & _
          " (CASE  " & _
          " WHEN REGEXP_LIKE(TRIM(MAX(a.WEIGHTOUT)), '^\d+(\.\d+)?$') THEN TO_NUMBER(MAX(a.WEIGHTOUT))  " & _
          " ELSE 0  " & _
          " END) AS TBSSUSUTKG, " & _
          " ROUND(((CASE WHEN REGEXP_LIKE(TRIM(MAX(f.QUANTITY)), '^\d+(\.\d+)?$') THEN TO_NUMBER(MAX(f.QUANTITY)) ELSE 0 END) -  " & _
          " (CASE WHEN REGEXP_LIKE(TRIM(MAX(a.WEIGHTOUT)), '^\d+(\.\d+)?$') THEN TO_NUMBER(MAX(a.WEIGHTOUT)) ELSE 0 END)) /  " & _
          " NULLIF(CASE WHEN REGEXP_LIKE(TRIM(MAX(a.WEIGHTOUT)), '^\d+(\.\d+)?$') THEN TO_NUMBER(MAX(a.WEIGHTOUT)) ELSE NULL END, 0), 2) AS TBSSUSUTPER, " & _
          " MAX(e.REMARKS) AS NOTE, " & _
          " NVL(MAX(a.PRICE_TENGKULAK), 0) AS HRGTENGKULAK, " & _
          " '0' AS DIBAYAR, " & _
          " NVL(MAX(a.PRICE_TBS), 0) AS HARGATBS, " & _
          " '0' AS GROSS, " & _
          " '0' AS PPH, " & _
          " '0' AS RATEPPH, " & _
          " NVL(MAX(a.RATEPPN), 0) AS RATEPPN, " & _
          " '0' AS REVENUE " & _
          " FROM LHPTIKETHDR a " & _
          "  LEFT JOIN ( " & _
          "      SELECT DOCKETID, SUM(TANDANQTY) AS TANDANQTY  " & _
          "      FROM LHPFFBWBTRN  " & _
          "      GROUP BY DOCKETID " & _
          "  ) b ON a.DOCKETID = b.DOCKETID " & _
          "     LEFT JOIN ( " & _
          "         SELECT DOCKETID, SUM(TANDANQTY) AS TANDANQTY  " & _
          "         FROM LHPFFBWBTRN  " & _
          "         GROUP BY DOCKETID " & _
          "     ) b ON a.DOCKETID = b.DOCKETID " & _
          "     LEFT JOIN LHPPKSMST c ON TO_CHAR(a.IDPKSMST) = c.EPLASCODE " & _
          "     LEFT JOIN SALESSHIPMENTDTL f ON a.DOCKETID = f.WBT_CHITNO " & _
          "     LEFT JOIN SALESSHIPMENTHDR e ON f.SALESSHIPMENTNO = e.SALESSHIPMENTNO " & _
          "     WHERE a.FCBA = '" & Request("fcba") & "' " & _
          "     AND a.DOCKETDATE BETWEEN TO_DATE('" & tglfrom & "', 'MM/DD/YYYY') AND TO_DATE('" & tglto   & "', 'MM/DD/YYYY') " & _
          "     GROUP BY a.FCBA, a.DOCKETID, a.DOCKETDATE, a.SPBNO, a.NOREG, a.DRIVERNAME, b.TANDANQTY " & _
          " ) res order by res.TGLTIMBANG ASC"
    ' Response.write sql
    Set setdata = Server.CreateObject("ADODB.Recordset")
    setdata.Open sql, connStr, adOpenStatic

    Dim totpksin, totpksout, totNetto1, totPot, totNetto2
    totJanjang = 0 : totTarra  = 0 : totBruto  = 0 : totNetto  = 0
    totpksin   = 0 : totpksout = 0 : totNetto1 = 0 : totPot    = 0 : totNetto2 = 0

    Function SafeDbl(v)
      If IsNull(v) Or IsEmpty(v) Or Trim(v & "") = "" Then
        SafeDbl = 0
      ElseIf IsNumeric(v) Then
        SafeDbl = CDbl(v)
      Else
        SafeDbl = 0
      End If
    End Function

    Function FmtN(v, dec)
      Dim d
      d = SafeDbl(v)
      If d > 2147483647 Or d < -2147483647 Then
        FmtN = CStr(CLng(d))
      Else
        FmtN = FormatNumber(d, dec)
      End If
    End Function

    Function AmbilJam(sDateTime)
      Dim s : s = Trim(sDateTime & "")
      If s = "" Then AmbilJam = "" : Exit Function
      If InStr(s, " ") > 0 Then
        Dim parts : parts = Split(s, " ")
        Dim jamStr : jamStr = parts(1)
        If Len(jamStr) > 5 Then jamStr = Left(jamStr, 5)
        If UBound(parts) >= 2 Then
          AmbilJam = jamStr & " " & parts(2)
        Else
          AmbilJam = jamStr
        End If
      Else
        If Len(s) > 5 Then AmbilJam = Left(s, 5) Else AmbilJam = s
      End If
    End Function
  %>

  <div class="wrap">
    <table id="tblMain" >
      <colgroup>
        <col style="width:24px">
        <col style="width:34px">
        <col style="width:52px">
        <col style="width:62px">
        <col style="width:66px">
        <col style="width:60px">
        <col style="width:130px">
        <col style="width:50px">
        <col style="width:50px">
        <col style="width:64px">
        <col style="width:72px">
        <col style="width:52px">
        <col style="width:58px">
        <col style="width:60px">
        <col style="width:92px">
        <col style="width:58px">
        <col style="width:58px">
        <col style="width:64px">
        <col style="width:64px">
        <col style="width:62px">
        <col style="width:50px">
        <col style="width:50px">
        <col style="width:62px">
        <col style="width:64px">
        <col style="width:68px">
        <col style="width:64px">
        <col style="width:64px">
        <col style="width:68px">
        <col style="width:58px">
        <col style="width:64px">
        <col style="width:58px">
        <col style="width:40px">
        <col style="width:80px">
        <col style="width:84px">
        <col style="width:80px">
        <col style="width:84px">
        <col style="width:60px">
        <col style="width:52px">
        <col style="width:72px">
        <col style="width:84px">
      </colgroup>

      <thead>
        <tr>
          <th rowspan="2" class="gh-no">No.</th>
          <th colspan="3" style="padding: 10px">Kebun</th>
          <th rowspan="2" style="padding: 10px">Tujuan</th>
          <th colspan="9" style="padding: 10px">Timbangan Kebun</th>
          <th colspan="11"style="padding: 10px">Timbangan PKS</th>
          <th colspan="7" style="padding: 10px">Data TBS Dipulangkan</th>
          <th colspan="9" style="padding: 10px">Inputan Commercial HO</th>
        </tr>
        <tr>
          <th class="sh-kb">Kebun</th>
          <th class="sh-kb">Janjang</th>
          <th class="sh-kb">Tonase</th>
          <th class="sh-tk">Tgl.Timbang</th>
          <th class="sh-tk">No. SPB</th>
          <th class="sh-tk">Jam Masuk</th>
          <th class="sh-tk">Jam Keluar</th>
          <th class="sh-tk">No. Polisi</th>
          <th class="sh-tk">Nama Supir</th>
          <th class="sh-tk">Tara (kg)</th>
          <th class="sh-tk">Bruto (kg)</th>
          <th class="sh-tk">Netto (kg)</th>
          <th class="sh-pk">No. Tiket</th>
          <th class="sh-pk">Tgl Masuk</th>
          <th class="sh-pk">Tgl Keluar</th>
          <th class="sh-pk">Masuk (kg)</th>
          <th class="sh-pk">Keluar (kg)</th>
          <th class="sh-pk">Netto 1 PKS (kg)</th>
          <th class="sh-pk">Pot (%)</th>
          <th class="sh-pk">Pot (kg)</th>
          <th class="sh-pk">Netto 2 (kg)</th>
          <th class="sh-pk">No. Polisi</th>
          <th class="sh-pk">Driver</th>
          <th class="sh-dp">Dipulangkan (kg)</th>
          <th class="sh-dp">Netto Tengkulak (kg)</th>
          <th class="sh-dp">Selisih x Tengkulak (kg)</th>
          <th class="sh-dp">Dipulangkan (%)</th>
          <th class="sh-dp">Susut (kg)</th>
          <th class="sh-dp">Susut (%)</th>
          <th class="sh-dp">Note</th>
          <th class="sh-cm">Tengkulak (Rp)</th>
          <th class="sh-cm">Dibayar (Rp)</th>
          <th class="sh-cm">Harga TBS (Rp)</th>
          <th class="sh-cm">Gross Revenue (Rp)</th>
          <th class="sh-cm">PPH (0.25%)</th>
          <th class="sh-cm">Rate PPN</th>
          <th class="sh-cm">PPN (Rp)</th>
          <th class="sh-cm">Revenue (Rp)</th>
        </tr>
      </thead>

      <tbody>
      <%
        If setdata.EOF Then
      %>
        <tr>
          <td colspan="41" style="text-align:center;padding:12px;color:#888;font-style:italic">
            Tidak ada data
          </td>
        </tr>
      <%
        Else
          Dim vNo, vFcba, vDocketId, vJanjang, vTonase, vPksTujuan
          Dim vTglTimbang, vSpbNo, vJamMasuk, vJamKeluar
          Dim vNopol, vSupir, vTarra, vBruto, vNetto
          Dim vNotiket, vTglMasukPks, vTglOutPks
          Dim vBeratIn, vBeratOut, vNettoPks, vPotPer, vPotKg, vNettoPks2
          Dim vNopolPks, vDriverPks
          Dim vTbsPulang, vNettoTeng, vSelisih, vTbsPer, vSusutKg, vSusutPer, vNote
          Dim vHrgTeng, vDibayar, vHrgTbs, vGross, vPph, vRatePpn, vPpn, vRevenue
          Dim kbClass
          Dim dblTbsPulang, dblNetto, dblSelisih, dblAC, tbsplng, susutkg, susutper
          Dim initHrgTeng, initHrgTbs, totJanjang, totTarra, totBruto, totNetto
          Dim dipulangkan, nettoTgk, SelisihTgk, TbsSusutKg
          Dim totDipulangkanPer, totSusutPer, totRows
          Dim dblNettoTeng

          totDipulangkanPer = 0
          totSusutPer       = 0
          totRows           = 0

          Do While Not setdata.EOF
            vNo         = setdata("no")
            vFcba       = setdata("FCBA")            & ""
            vDocketId   = setdata("DOCKETID")        & ""
            vJanjang    = setdata("TANDANQTY")        & ""
            vTonase     = setdata("estimasitonase")   & ""
            vPksTujuan  = setdata("pkstujuan")        & ""
            vTglTimbang = setdata("tgltimbang")       & ""
            vSpbNo      = setdata("SPBNO")            & ""
            vJamMasuk   = setdata("jammasuk")         & ""
            vJamKeluar  = setdata("jamkeluar")        & ""
            vNopol      = setdata("nopol")            & ""
            vSupir      = setdata("namasupir")        & ""
            vTarra      = setdata("tarra")            
            vBruto      = setdata("bruto")            
            vNetto      = setdata("netto")            
            vNotiket     = setdata("notiket")         & ""
            vTglMasukPks = setdata("tglmasukpks")     & ""
            vTglOutPks   = setdata("tgloutpks")       & ""
            vBeratIn     = setdata("beratinpks")      & ""
            vBeratOut    = setdata("beratoutpks")     & ""
            vNettoPks    = setdata("nettopks")        & ""
            vPotPer      = setdata("potpksper")       & ""
            vPotKg       = setdata("potpkskg")        & ""
            vNettoPks2   = setdata("nettopks2")       & ""
            vNopolPks    = setdata("nopolpks")        & ""
            vDriverPks   = setdata("driverpks")       & ""
            vTbsPulang   = setdata("tbspulang")       & ""
            vNettoTeng   = setdata("nettotengkulak")  & ""
            vSelisih     = setdata("selisih")         & ""
            vTbsPer      = setdata("tbsbalikper")     & ""
            vSusutKg     = setdata("tbssusutkg")      & ""
            vSusutPer    = setdata("tbssusutper")     & ""
            vNote        = setdata("note")            & ""
            vHrgTeng     = setdata("hrgtengkulak")    & ""
            vHrgTbs      = setdata("hargatbs")        & ""
            vRatePpn     = setdata("RATEPPN") & ""

            If vRatePpn = "" Then vRatePpn = "0"
            If vTbsPulang = "" Or IsNull(vTbsPulang) Then dblTbsPulang = 0 Else dblTbsPulang = CDbl(vTbsPulang)
            If vNetto     = "" Or IsNull(vNetto)     Then dblNetto     = 0 Else dblNetto     = CDbl(vNetto)

            If dblNetto <> 0 Then dblAC = dblTbsPulang / dblNetto Else dblAC = 0

            Dim dblBeratOut, dblTarra2, dblBeratIn, dblBruto2

            dblBeratOut = 0 : dblTarra2  = 0
            dblBeratIn  = 0 : dblBruto2  = 0

            totDipulangkanPer = totDipulangkanPer + dblAC     ' dblAC = tbsplng / dblNetto
            totSusutPer       = totSusutPer       + susutper  ' susutper = (beratIn - bruto) / bruto
            totRows           = totRows + 1

            If Not IsNull(vBeratOut) And vBeratOut <> "" Then dblBeratOut = CDbl(vBeratOut)
            If Not IsNull(vTarra)    And vTarra    <> "" Then dblTarra2   = CDbl(vTarra)
            If Not IsNull(vBeratIn)  And vBeratIn  <> "" Then dblBeratIn  = CDbl(vBeratIn)
            If Not IsNull(vBruto)    And vBruto    <> "" Then dblBruto2   = CDbl(vBruto)

            tbsplng  = dblBeratOut - dblTarra2
            susutkg  = dblBeratIn  - dblBruto2
            If dblBruto2 <> 0 Then
                susutper = (dblBeratIn - dblBruto2) / dblBruto2
            Else
                susutper = 0
            End If

            kbClass = ""
            Select Case UCase(vFcba)
              Case "CDE" : kbClass = "cde"
              Case "BIA" : kbClass = "bia"
              Case "GHE" : kbClass = "ghe"
            End Select

            If vJanjang <> "" Then totJanjang = totJanjang + CDbl(vJanjang)
            totTarra = totTarra + CDbl(vTarra)
            totBruto = totBruto + CDbl(vBruto)
            totNetto = totNetto + CDbl(vNetto)
            totNetto   = totNetto   + SafeDbl(vNetto)
            totpksin   = totpksin   + SafeDbl(vBeratIn)
            totpksout  = totpksout  + SafeDbl(vBeratOut)
            totNetto1  = totNetto1  + SafeDbl(vNettoPks)
            totPot     = totPot     + SafeDbl(vPotKg)
            totNetto2  = totNetto2  + SafeDbl(dblNettoPks2)
            dipulangkan  = dipulangkan  + SafeDbl(tbsplng)
            nettoTgk  = nettoTgk  + SafeDbl(vNettoTeng)
            SelisihTgk = SelisihTgk + SafeDbl(CDbl(vNettoTeng) - tbsplng)
            TbsSusutKg  = dipulangkan  + SafeDbl(dblAC)

            initHrgTeng = ""
            initHrgTbs  = ""
            If vHrgTeng <> "0" And vHrgTeng <> "" Then initHrgTeng = vHrgTeng
            If vHrgTbs  <> "0" And vHrgTbs  <> "" Then initHrgTbs  = vHrgTbs
      
            Dim sPpn0, sPpn1, sPpn2, sTouched
            Dim dblRatePpn
            If vRatePpn = "" Then
              dblRatePpn = 0
            Else
              dblRatePpn = CDbl(vRatePpn)
            End If

            sPpn0   = "" : sPpn1   = "" : sPpn2   = ""
            sTouched = "false"

            If dblRatePpn = 0 Then
              sPpn0 = "selected"
            ElseIf dblRatePpn > 0.01 And dblRatePpn < 0.02 Then
              sPpn1    = "selected"
              sTouched = "true"
            ElseIf dblRatePpn > 0.1 Then
              sPpn2    = "selected"
              sTouched = "true"
            End If

            Dim dblPotKg, dblNettoPks, dblPotPer, dblNettoPks2
            dblPotKg    = 0 : dblNettoPks  = 0
            dblPotPer   = 0 : dblNettoPks2 = 0

            If Not IsNull(setdata("potpkskg"))  And setdata("potpkskg")  <> "" Then dblPotKg    = CDbl(setdata("potpkskg"))
            If Not IsNull(setdata("nettopks"))  And setdata("nettopks")  <> "" Then dblNettoPks  = CDbl(setdata("nettopks"))
            If Not IsNull(setdata("beratinpks")) And setdata("beratinpks") <> "" Then dblBeratIn  = CDbl(setdata("beratinpks"))
            If Not IsNull(setdata("beratoutpks")) And setdata("beratoutpks") <> "" Then dblBeratOut = CDbl(setdata("beratoutpks"))

            If dblNettoPks <> 0 Then dblPotPer   = (dblPotKg / dblNettoPks) * 100
            dblNettoPks2 = dblNettoPks - dblPotKg
          %>
          <tr>
            <td class="td-no"><%=vNo%></td>
            <td class="td-kb"><span class="<%=kbClass%>"><%=vFcba%></span></td>
            <td class="td-r"><%=FormatNumber(vJanjang,0)%></td>
            <td class="td-r"><%=FormatNumber(vNetto,0)%></td>
            <td><%=vPksTujuan%></td>
            <td><%=vTglTimbang%></td>
            <td class="td-l"><%=vSpbNo%></td>
            <td><%=AmbilJam(vJamMasuk)%></td>
            <td><%=AmbilJam(vJamKeluar)%></td>
            <td class="td-l"><%=vNopol%></td>
            <td class="td-l"><%=vSupir%></td>
            <td class="td-r"><%=FormatNumber(vTarra,0)%></td>
            <td class="td-r"><%=FormatNumber(vBruto,0)%></td>
            <td class="td-cy"><%=FormatNumber(vNetto,0)%></td>
            <td class="td-l"><%=vNotiket%></td>
            <td><%=vTglMasukPks%></td>
            <td><%=vTglOutPks%></td>
            <td class="td-r"><%=FormatNumber(vBeratIn,0)%></td>
            <td class="td-r"><%=FormatNumber(vBeratOut,0)%></td>
            <td class="td-cy"><%=FormatNumber(vNettoPks,0)%></td>
            <td class="td-cy"><%=FormatNumber(dblPotPer, 2)%> %</td>
            <td class="td-r"><%=vPotKg%></td>
            <td class="td-cy"><%=FormatNumber(dblNettoPks2, 0)%></td>
            <td class="td-l"><%=vNopolPks%></td>
            <td class="td-l"><%=vDriverPks%></td>
            <td class="td-cy"><%=FormatNumber(tbsplng,0)%></td>
            <td class="td-r"><%=FormatNumber(vNettoTeng,0)%></td>
            <td class="td-cy"><%=FormatNumber(vNettoTeng-tbsplng,0)%></td>
            <td class="td-cy"><%=FormatNumber(dblAC,2)%>%</td>
            <td class="td-cy"><%=FormatNumber(susutkg,0)%></td>
            <td class="td-cy"><%=FormatNumber(susutper,2)%>%</td>
            <td><%=vNote%></td>

            <!-- index 32 : Harga Tengkulak INPUT -->
            <td class="td-r" style="padding:1px 2px">
              <input type="hidden" class="h-fcba"     value="<%=vFcba%>">
              <input type="hidden" class="h-docketid" value="<%=vDocketId%>">
              <input class="inp-rp" type="number" min="0" step="1"
                placeholder="0"
                value="<%=initHrgTeng%>"
                data-col="hrg-teng"
                data-field="PRICE_TENGKULAK"
                oninput="calcRow(this.closest('tr'))"
                onchange="fillAllCol(this); autoSave(this)">
            </td>

            <!-- index 33 : AH kalkulasi -->
            <td class="td-cy" data-calc="AH">0</td>

            <!-- index 34 : Harga TBS INPUT -->
            <td class="td-r" style="padding:1px 2px">
              <input class="inp-rp" type="number" min="0" step="1"
                placeholder="0"
                value="<%=initHrgTbs%>"
                data-col="hrg-tbs"
                data-field="PRICE_TBS"
                oninput="calcRow(this.closest('tr'))"
                onchange="fillAllCol(this); autoSave(this)">
            </td>

            <!-- index 35-39 : kalkulasi -->
            <td class="td-cy" data-calc="AJ">0</td>
            <td class="td-cy" data-calc="AK">0</td>
            <td class="td-cy">
              <select class="sel-ppn"
                data-col="rate-ppn"
                data-field="RATEPPN"
                data-touched="<%=sTouched%>"
                onchange="this.setAttribute('data-touched','true'); calcRow(this.closest('tr')); fillAllCol(this); autoSaveSelect(this)"
                style="width:100%;font-size:12px;font-family:Arial,sans-serif;
                      background:#fff9c4;border:1px solid #aaa;border-radius:2px;
                      padding:1px 2px;cursor:pointer">
                <option value="0"     <%=sPpn0%>>0%</option>
                <option value="0.011" <%=sPpn1%>>1,1%</option>
                <option value="0.11"  <%=sPpn2%>>11%</option>
              </select>
            </td>
            <td class="td-cy" data-calc="AL">0</td>
            <td class="td-cy" data-calc="AM">0</td>
          </tr>
      <%
            setdata.MoveNext
          Loop
        End If

        setdata.Close
        Set setdata = Nothing

        Dim avgDipulangkanPer, avgSusutPer
        If totRows > 0 Then
            avgDipulangkanPer = totDipulangkanPer / totRows
            avgSusutPer       = totSusutPer / totRows
        Else
            avgDipulangkanPer = 0
            avgSusutPer       = 0
        End If
      %>
      </tbody>

      <tfoot>
        <tr class="tot">
          <td colspan="2" style="text-align:center;font-weight:bold">TOTAL</td>
          <td class="td-r"><%=FormatNumber(totJanjang,0)%></td>
          <td style="text-align: right"><%=FormatNumber(totNetto,0)%></td>
          <td colspan="7"></td>
          <td class="td-r"><%=FormatNumber(totTarra,0)%></td>
          <td class="td-r"><%=FormatNumber(totBruto,0)%></td>
          <td class="td-r" id="totNetto"><%=FormatNumber(totNetto,0)%></td>
          <td colspan="3"></td>
          <td style="text-align: right"><%=FmtN(totpksin,0)%></td>
          <td style="text-align: right"><%=FmtN(totpksout,0)%></td>
          <td style="text-align: right"><%=FmtN(totNetto1,0)%></td>
          <td></td>
          <td style="text-align: right"><%=FmtN(totPot,0)%></td>
          <td style="text-align: right"><%=FmtN(totNetto2,0)%></td>
          <td colspan="2"></td>
          <td style="text-align: right"><%=FmtN(dipulangkan,0)%></td>
          <td style="text-align: right"><%=FmtN(nettoTgk,0)%></td>
          <td style="text-align: right"><%=FormatNumber(SelisihTgk,0)%></td>
          <td style="text-align:right"></td>
          <td style="text-align: right"><%=FmtN(TbsSusutKg,0)%></td>
          <td style="text-align:right"><%=FormatNumber(avgSusutPer, 2)%>%</td>
          <td></td>
          <td></td>
          <td id="totAH"></td>
          <td></td>
          <td id="totAJ"></td>
          <td id="totAK"></td>
          <td></td>
          <td id="totAL"></td>
          <td id="totAM"></td>
        </tr>
      </tfoot>
    </table>
  </div>
</div>

<script>
/* ════════════════════════════════════════════════════
   Nilai dari server
   ════════════════════════════════════════════════════ */
var companyName = "<%=setCompany("FCCOMPANYNAME")%>";
var companyCode = "<%=setCompany("FCCOMPANYCODE")%>";
var tahun       = "<%=Year(CDate(Request("datemontbsto")))%>";
var fcba        = "<%=Request("fcba")%>";

/* ════════════════════════════════════════════════════
   DOMContentLoaded
   ════════════════════════════════════════════════════ */
document.addEventListener('DOMContentLoaded', function () {
  /* Sticky header — hitung top secara akurat */
  var headerRows = document.querySelectorAll('thead tr');
  var tops = [];
  var cumTop = 0;

  headerRows.forEach(function (row, i) {
    tops[i] = cumTop;
    cumTop += row.offsetHeight;
  });

  headerRows.forEach(function (row, i) {
    row.querySelectorAll('th').forEach(function (th) {
      th.style.position = 'sticky';
      th.style.top      = tops[i] - 2 + 'px';
      th.style.zIndex   = '3';
      th.style.backgroundColor = th.style.backgroundColor || '#499e10';
    });
  });

  /* Kolom pertama thead dapat z-index lebih tinggi */
  document.querySelectorAll('thead tr th:first-child').forEach(function(th) {
    th.style.zIndex = '6';
  });

  /* Sticky kolom pertama tbody & tfoot */
  document.querySelectorAll('tbody tr td:first-child, tfoot tr td:first-child').forEach(function(td) {
    td.style.position = 'sticky';
    td.style.left     = '0';
    td.style.zIndex   = '2';
  });

  /* Hitung semua baris saat load */
  document.querySelectorAll('#tblMain tbody tr').forEach(function(tr) {
    calcRow(tr);
  });
  updateTotals();
});

/* ════════════════════════════════════════════════════
   Format & parse angka
   ════════════════════════════════════════════════════ */
function fmtN(n) {
  if (isNaN(n) || n === null || n === undefined) return '0';
  return Math.round(n).toLocaleString('id-ID');
}

function fmtRp(n) {
  if (isNaN(n) || n === null || n === undefined) return 'Rp 0';
  return Math.round(n).toLocaleString('id-ID');
}

function parseN(str) {
  if (str === null || str === undefined) return 0;
  var s = String(str).replace(/\./g, '').replace(',', '.');
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function parseNRp(str) {
  if (str === null || str === undefined) return 0;
  var s = String(str)
            .replace(/,/g, '')   // hapus semua koma (thousand separator en-US)
            .trim();
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

/* ════════════════════════════════════════════════════
   calcRow — hitung satu baris
   Peta td (0-based):
     0  No | 1 Kebun | 2 Janjang | 3 Tonase | 4 PKS Tujuan
     5-13  Timbangan Kebun (termasuk Netto di 13)
     14-24 Timbangan PKS
     25-31 TBS Dipulangkan
     32    INPUT Harga Tengkulak (AG) + hidden FCBA, DOCKETID
     33    AH (kalkulasi)
     34    INPUT Harga TBS (AI)
     35    AJ | 36 AK | 37 Rate PPN | 38 AL | 39 AM
     40    Status cell
   ════════════════════════════════════════════════════ */
function calcRow(tr) {
  var cells = tr.querySelectorAll('td');
  if (cells.length < 40) return;

  var AA = parseNRp(cells[26].innerText);

  var inpAG = cells[32].querySelector('input.inp-rp');
  var AG    = inpAG ? (parseFloat(inpAG.value) || 0) : 0;

  var inpAI = cells[34].querySelector('input.inp-rp');
  var AI    = inpAI ? (parseFloat(inpAI.value) || 0) : 0;

  /* Baca rate PPN dari dropdown — cells[37] */
  var selPPN  = cells[37].querySelector('select.sel-ppn');
  var ratePPN = selPPN ? (parseFloat(selPPN.value) || 0) : 0;

  var AH = AA * AG;
  var AJ = AI * AA;
  var AK = AJ * 0.0025;
  var AL = AJ * ratePPN;   /* pakai rate dari dropdown */
  var AM = AJ - AK + AL;

  cells[33].innerText = fmtRp(AH);
  cells[35].innerText = fmtRp(AJ);
  cells[36].innerText = fmtRp(AK);
  cells[38].innerText = fmtRp(AL);
  cells[39].innerText = fmtRp(AM);

  updateTotals();
}

/* ════════════════════════════════════════════════════
   updateTotals
   ════════════════════════════════════════════════════ */
function updateTotals() {
  var sumAH = 0, sumAJ = 0, sumAK = 0, sumAL = 0, sumAM = 0;
  document.querySelectorAll('#tblMain tbody tr').forEach(function(tr) {
    var cells = tr.querySelectorAll('td');
    if (cells.length < 40) return;
    sumAH += parseN(cells[33].innerText);
    sumAJ += parseN(cells[35].innerText);
    sumAK += parseN(cells[36].innerText);
    sumAL += parseN(cells[38].innerText);
    sumAM += parseN(cells[39].innerText);
  });
  var el = function(id){ return document.getElementById(id); };
  if (el('totAH')) { el('totAH').innerText = fmtRp(sumAH); el('totAH').style.textAlign = 'right'; }
  if (el('totAJ')) { el('totAJ').innerText = fmtRp(sumAJ); el('totAJ').style.textAlign = 'right'; }
  if (el('totAK')) { el('totAK').innerText = fmtRp(sumAK); el('totAK').style.textAlign = 'right'; }
  if (el('totAL')) { el('totAL').innerText = fmtN(sumAL); el('totAL').style.textAlign = 'right'; }
  if (el('totAM')) { el('totAM').innerText = fmtN(sumAM); el('totAM').style.textAlign = 'right'; }
}

/* ════════════════════════════════════════════════════
   fillAllCol — broadcast nilai ke semua baris kolom
   Hanya mengisi baris yang masih kosong / 0.
   ════════════════════════════════════════════════════ */
function fillAllCol(inp) {
  var col = inp.getAttribute('data-col');
  var val = inp.value;
  if (val === '' || val === null) return;

  var isSelect = inp.tagName === 'SELECT';

  var selector = isSelect
    ? '#tblMain tbody select[data-col="' + col + '"]'
    : '#tblMain tbody input.inp-rp[data-col="' + col + '"]';

  document.querySelectorAll(selector).forEach(function(other) {
    if (other === inp) return;

    if (!isSelect) {
      if (other.value !== '' && other.value !== '0') return;
      other.value = val;
      other.classList.add('filled-auto');
      setTimeout(function(){ other.classList.remove('filled-auto'); }, 1500);
      autoSave(other);
    } else {
      if (other.getAttribute('data-touched') === 'true') return;
      other.value = val;
      other.setAttribute('data-touched', 'true');  /* <-- tambahkan ini */
      other.style.background = '#d4edda';
      setTimeout(function(){ other.style.background = '#fff9c4'; }, 1500);
      autoSaveSelect(other);
    }

    calcRow(other.closest('tr'));
  });

  updateTotals();
}

/* ════════════════════════════════════════════════════
   autoSave — simpan PRICE_TENGKULAK / PRICE_TBS ke DB
   via AJAX POST ke halaman ini sendiri (self-POST)
   ════════════════════════════════════════════════════ */
var _saveTimers = {};

function autoSave(inp) {
  var tr       = inp.closest('tr');
  var cells    = tr.querySelectorAll('td');
  var cell32   = cells[32];

  var docketId = cell32.querySelector('input.h-docketid');
  var fcbaVal  = cell32.querySelector('input.h-fcba');

  if (!docketId || !fcbaVal) {
    showStatus('DOCKETID / FCBA tidak ditemukan di baris ini.', 'err');
    return;
  }

  var field    = inp.getAttribute('data-field');
  var value    = parseFloat(inp.value) || 0;
  var statusTd = tr.querySelector('[data-status-cell]');
  var key      = fcbaVal.value + '_' + docketId.value + '_' + field;

  if (_saveTimers[key]) clearTimeout(_saveTimers[key]);
  _saveTimers[key] = setTimeout(function() {

    inp.classList.remove('state-saved', 'state-error', 'filled-auto');
    inp.classList.add('state-saving');
    if (statusTd) statusTd.innerText = '⏳';

    var xhr = new XMLHttpRequest();
    /* POST ke halaman ini sendiri */
    xhr.open('POST', window.location.pathname, true);
    xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');

    xhr.onreadystatechange = function() {
      if (xhr.readyState !== 4) return;

      inp.classList.remove('state-saving');

      if (xhr.status === 200 && xhr.responseText.trim() === 'OK') {
        inp.classList.add('state-saved');
        if (statusTd) statusTd.innerText = '✔';
        showStatus('Tersimpan — DOCKETID ' + docketId.value + ' [' + field + ']', 'ok');
        setTimeout(function(){ inp.classList.remove('state-saved'); }, 3000);
      } else {
        inp.classList.add('state-error');
        if (statusTd) statusTd.innerText = '✘';
        var errMsg = xhr.responseText.trim() || ('HTTP ' + xhr.status);
        showStatus('Gagal simpan [' + field + ']: ' + errMsg, 'err');
      }
    };

    xhr.onerror = function() {
      inp.classList.remove('state-saving');
      inp.classList.add('state-error');
      if (statusTd) statusTd.innerText = '✘';
      showStatus('Koneksi error saat menyimpan ' + field, 'err');
    };

    var body = 'fcba='      + encodeURIComponent(fcbaVal.value)  +
               '&docketid=' + encodeURIComponent(docketId.value) +
               '&field='    + encodeURIComponent(field)          +
               '&value='    + encodeURIComponent(value);
    xhr.send(body);

  }, 600);
}

/* ════════════════════════════════════════════════════
   autoSaveSelect — simpan nilai dropdown (RATEPPN, dll)
   via AJAX POST ke halaman ini sendiri (self-POST)
   ════════════════════════════════════════════════════ */
function autoSaveSelect(sel) {
  var tr      = sel.closest('tr');
  var cells   = tr.querySelectorAll('td');
  var cell32  = cells[32];

  var docketId = cell32.querySelector('input.h-docketid');
  var fcbaVal  = cell32.querySelector('input.h-fcba');

  if (!docketId || !fcbaVal) {
    showStatus('DOCKETID / FCBA tidak ditemukan di baris ini.', 'err');
    return;
  }

  var field    = sel.getAttribute('data-field');
  var value    = parseFloat(sel.value) || 0;
  var statusTd = tr.querySelector('[data-status-cell]');
  var key      = fcbaVal.value + '_' + docketId.value + '_' + field;

  if (_saveTimers[key]) clearTimeout(_saveTimers[key]);
  _saveTimers[key] = setTimeout(function() {

    sel.style.borderColor = '#f9a825';
    sel.style.background  = '#fff9c4';
    if (statusTd) statusTd.innerText = '⏳';

    var xhr = new XMLHttpRequest();
    xhr.open('POST', window.location.pathname, true);
    xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');

    xhr.onreadystatechange = function() {
      if (xhr.readyState !== 4) return;

      if (xhr.status === 200 && xhr.responseText.trim() === 'OK') {
        sel.style.borderColor = '#388e3c';
        sel.style.background  = '#e8f5e9';
        if (statusTd) statusTd.innerText = '✔';
        showStatus('Tersimpan — DOCKETID ' + docketId.value + ' [' + field + ']', 'ok');
        setTimeout(function(){
          sel.style.borderColor = '#aaa';
          sel.style.background  = '#fff9c4';
        }, 3000);
      } else {
        sel.style.borderColor = '#c62828';
        sel.style.background  = '#ffebee';
        if (statusTd) statusTd.innerText = '✘';
        var errMsg = xhr.responseText.trim() || ('HTTP ' + xhr.status);
        showStatus('Gagal simpan [' + field + ']: ' + errMsg, 'err');
      }
    };

    xhr.onerror = function() {
      sel.style.borderColor = '#c62828';
      sel.style.background  = '#ffebee';
      if (statusTd) statusTd.innerText = '✘';
      showStatus('Koneksi error saat menyimpan ' + field, 'err');
    };

    var body = 'fcba='      + encodeURIComponent(fcbaVal.value)  +
               '&docketid=' + encodeURIComponent(docketId.value) +
               '&field='    + encodeURIComponent(field)          +
               '&value='    + encodeURIComponent(value);
    xhr.send(body);

  }, 600);
}

/* ════════════════════════════════════════════════════
   showStatus
   ════════════════════════════════════════════════════ */
function showStatus(msg, type) {
  var bar = document.getElementById('statusBar');
  var el  = document.createElement('div');
  el.className  = 'status-msg ' + (type || 'inf');
  el.innerText  = msg;
  bar.appendChild(el);
  setTimeout(function(){ el.style.opacity = '0'; }, 3000);
  setTimeout(function(){ if (el.parentNode) el.parentNode.removeChild(el); }, 4000);
}

/* ════════════════════════════════════════════════════
   downloadExcel
   ════════════════════════════════════════════════════ */
function downloadExcel() {
  var inputs = document.querySelectorAll('#tblMain tbody input.inp-rp');
  inputs.forEach(function(inp) {
    var sp = document.createElement('span');
    sp.className = '__exp';
    sp.innerText = inp.value || '0';
    inp.style.display = 'none';
    inp.parentNode.appendChild(sp);
  });

  var statusCells = document.querySelectorAll('[data-status-cell]');
  statusCells.forEach(function(td){ td.style.display = 'none'; });

  var wb      = XLSX.utils.book_new();
  var ws_data = [];
  ws_data.push(["FIRST BORNEO GROUP"]);
  ws_data.push(["PT. " + companyName]);
  ws_data.push([]);
  ws_data.push([companyCode + " / " + tahun, "MONITORING PRODUKSI TBS PER MOBIL PER HARI"]);
  ws_data.push([]);

  var tbl     = document.querySelector('#tblMain');
  var tblData = XLSX.utils.sheet_to_json(
    XLSX.utils.table_to_sheet(tbl, {raw: false}),
    {header: 1}
  );

  var allData = ws_data.concat(tblData);
  var ws      = XLSX.utils.aoa_to_sheet(allData);
  ws['!cols'] = [{wch:12},{wch:50}];
  XLSX.utils.book_append_sheet(wb, ws, "Monitoring TBS");

  var now = new Date();
  var tgl = now.getFullYear() + '-' +
            String(now.getMonth()+1).padStart(2,'0') + '-' +
            String(now.getDate()).padStart(2,'0');
  XLSX.writeFile(wb, 'Monitoring_TBS_' + tgl + '.xlsx');

  inputs.forEach(function(inp) {
    inp.style.display = '';
    var sp = inp.parentNode.querySelector('span.__exp');
    if (sp) inp.parentNode.removeChild(sp);
  });
  statusCells.forEach(function(td){ td.style.display = ''; });
}
</script>
</body>
</html>
