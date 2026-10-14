<%Option Explicit%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="adovbs.inc"-->

<html>
<head>
<!--#include file="aspupload.asp"-->
<title>ePlaS - Upload Manual Fingerprint</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js"></script>
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
<script type="text/javascript" src="dtree.js"></script>
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
<script language="JavaScript">
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window);
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);
</script>
<script>
function onSubmitForm() {
    var formDOMObj = document.frmSend;
    if (formDOMObj.txtFile.value == "" && formDOMObj.fcCode.value == "(NEW)") {
        alert("Please press the Browse button and pick a file.");
        return false;
    }
    var fileName = formDOMObj.txtFile.value;
    var ext = fileName.substring(fileName.lastIndexOf('.')+1).toLowerCase();
    if (ext !== 'xls' && ext !== 'xlsx') {
        alert("File harus berformat .xls atau .xlsx");
        return false;
    }
    return true;
}
</script>
<script>
$(function() {
    $(".datepicker").datepicker();
});
</script>
<style>
.msg-box        { padding:8px 12px; margin:6px 0; border-radius:4px; font-family:Arial,sans-serif; font-size:13px; }
.msg-success    { background:#d4edda; color:#155724; border:1px solid #c3e6cb; }
.msg-error      { background:#f8d7da; color:#721c24; border:1px solid #f5c6cb; }
.msg-warning    { background:#fff3cd; color:#856404; border:1px solid #ffeeba; }
.msg-info       { background:#d1ecf1; color:#0c5460; border:1px solid #bee5eb; }
.tbl-result     { border-collapse:collapse; width:100%; margin-top:10px; font-size:12px; }
.tbl-result th  { background:#0C5B24; color:#fff; padding:5px 8px; text-align:left; }
.tbl-result td  { padding:4px 8px; border-bottom:1px solid #ddd; }
.tbl-result tr:nth-child(even) td { background:#f5f5f5; }
</style>
<% if Request("oldui") <> "1" then %>
<!-- Tampilan baru (css/fbp-shell.css + css/fbp-forms.css + css/fbp-docdetail.css). ?oldui=1 = tampilan lama -->
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<link rel="StyleSheet" href="css/login/font-awesome.min.css" type="text/css" />
<link rel="StyleSheet" href="css/fbp-shell.css?v=20261009d" type="text/css" />
<link rel="StyleSheet" href="css/fbp-forms.css?v=20261009e" type="text/css" />
<link rel="StyleSheet" href="css/fbp-docdetail.css?v=20261009b" type="text/css" />
<script type="text/javascript" src="js/fbp-shell.js?v=20261014a"></script>
<% end if %>
</head>
<body<% if Request("oldui") <> "1" then %> class="dash page-show page-legacy page-doc"<% end if %>>
<!--#include file="top.jsp"-->
<div style="width:21.7%; float:left; margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
    <!--#include file="menu.asp"-->
</div>
<div class="judul_head" valign="top" bgcolor="#fff" style="width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;">
    <div class="judul">
        Module : <%=Request("txtModule")%> / <%=Request("txtTable")%>
    </div>
    <div class="box_both" style="overflow:auto;">
<%
'=============================================================
' DEKLARASI VARIABEL
'=============================================================
Dim fcCode, cmdView, uploaddocdate, temptxtJenis
Dim Remarks, fcfilename, namefile
Dim txtModule, txtTable, txtSearch, txtField, txtYear, txtMonth
Dim sql, q, no
Dim error_flag, error_desc
Dim AutoNum
Dim uploadsDirVar
Dim stopProcess : stopProcess = False
Dim pageMsg     : pageMsg     = ""

'=============================================================
' FUNGSI Nz
'=============================================================
Function Nz(val, defaultVal)
    If IsNull(val) Or IsEmpty(val) Then
        Nz = defaultVal
    Else
        Nz = val
    End If
End Function

'=============================================================
' FUNGSI SafeSQL
'=============================================================
Function SafeSQL(strVal)
    If IsNull(strVal) Or IsEmpty(strVal) Then
        SafeSQL = ""
    Else
        SafeSQL = Replace(CStr(strVal), "'", "''")
    End If
End Function

'=============================================================
' FUNGSI IsValidExcelFile
'=============================================================
Function IsValidExcelFile(fileName)
    Dim ext
    ext = LCase(Mid(fileName, InStrRev(fileName, ".") + 1))
    IsValidExcelFile = (ext = "xls" Or ext = "xlsx")
End Function

'=============================================================
' FUNGSI GetEmpInfo
' Ambil fcba terbaru dari ATTENDANCE_TAB berdasarkan EMPNO
'=============================================================
Function GetEmpInfo(empno)
    Dim rsEmp, sqlEmp, result(0)
    result(0) = ""
    If Trim(empno) = "" Then
        GetEmpInfo = result
        Exit Function
    End If
    sqlEmp = " SELECT fcba"                          & _
             " FROM ("                               & _
             "   SELECT fcba"                        & _
             "   FROM userfinger.ATTENDANCE_TAB"     & _
             "   WHERE EMPNO = '" & SafeSQL(empno) & "'" & _
             "   ORDER BY timein DESC"               & _
             " ) WHERE ROWNUM = 1"
    Set rsEmp = Server.CreateObject("ADODB.Recordset")
    On Error Resume Next
    rsEmp.Open sqlEmp, connstr
    If Err.Number <> 0 Then
        Err.Clear
        Set rsEmp = Nothing
        GetEmpInfo = result
        Exit Function
    End If
    On Error GoTo 0
    If Not rsEmp.EOF Then
        result(0) = Nz(rsEmp("fcba"), "")
    End If
    rsEmp.Close
    Set rsEmp = Nothing
    GetEmpInfo = result
End Function

'=============================================================
' FUNGSI GenerateDocNo
' Generate DOCNO unik berdasarkan MAX nomor terakhir di tabel.
' Format: {FCBA}{YYYY}{MM}{NNN} — contoh: KNT202604001
' Menggunakan REGEXP_SUBSTR untuk ambil angka di ujung DOCNO
' sehingga tidak tergantung panjang prefix.
' Loop sampai dapat nomor yang benar-benar belum ada (max 50x).
'=============================================================
Function GenerateDocNo(fcbaVal, yearVal, monthVal)
    Dim rsLast, sqlLast
    Dim fcbaUp, yyStr, mmStr, lastSeq, newSeq
    Dim rsChk, sqlChk, docFound
    Dim tryCount, maxTry
    Dim dbErrDesc

    fcbaUp = UCase(Trim(fcbaVal))
    yyStr  = CStr(CInt(yearVal))
    mmStr  = Right("0" & CStr(CInt(monthVal)), 2)

    '--- DEBUG: cek fcba kosong ---
    If fcbaUp = "" Then
        pageMsg = pageMsg & "<div class='msg-box msg-error'>" & _
                  "GenerateDocNo: <b>fcBA kosong!</b> Pastikan Session(""fcBA"") terisi." & _
                  "</div>"
        GenerateDocNo = ""
        Exit Function
    End If

    '=====================================================
    ' LANGKAH 1: Ambil lastSeq via LIKE prefix DOCNO
    ' Lebih reliable daripada filter fcba+tdate
    '=====================================================
    Dim prefixDoc : prefixDoc = fcbaUp & yyStr & mmStr  ' contoh: KNT202604

    sqlLast = " SELECT NVL(MAX(TO_NUMBER(REGEXP_SUBSTR(docno, '[0-9]+$'))), 0) AS lastseq" & _
              " FROM userfinger.uploadheader_mf_tab"                                        & _
              " WHERE docno LIKE '" & SafeSQL(prefixDoc) & "%'"

    Set rsLast = Server.CreateObject("ADODB.Recordset")
    dbErrDesc  = ""
    On Error Resume Next
    rsLast.Open sqlLast, connStr   ' ← pakai connStr (huruf besar S), konsisten dengan Execute
    If Err.Number <> 0 Then
        lastSeq   = 0
        dbErrDesc = Err.Description
        Err.Clear
    Else
        If Not rsLast.EOF Then
            lastSeq = CInt(Nz(rsLast("lastseq"), 0))
        Else
            lastSeq = 0
        End If
    End If
    On Error GoTo 0
    If rsLast.State = 1 Then rsLast.Close
    Set rsLast = Nothing

    '--- DEBUG: tampilkan info lastSeq ---
    pageMsg = pageMsg & "<div class='msg-box msg-info'>" & _
              "GenerateDocNo &mdash; prefix: <b>" & prefixDoc & "</b> | " & _
              "lastSeq: <b>" & lastSeq & "</b>"
    If dbErrDesc <> "" Then
        pageMsg = pageMsg & " | <font color='red'>DB Error sqlLast: " & dbErrDesc & "</font>"
    End If
    pageMsg = pageMsg & "</div>"

    '=====================================================
    ' LANGKAH 2: Loop cari nomor yang belum ada
    '=====================================================
    tryCount = 0
    maxTry   = 450
    docFound = True

    Do While docFound And tryCount < maxTry
        tryCount = tryCount + 1
        newSeq   = lastSeq + tryCount
        AutoNum  = prefixDoc & Right("000" & CStr(newSeq), 3)

        sqlChk = " SELECT COUNT(1) cnt FROM userfinger.uploadheader_mf_tab" & _
                 " WHERE docno = '" & SafeSQL(AutoNum) & "'"

        Set rsChk = Server.CreateObject("ADODB.Recordset")
        dbErrDesc = ""
        On Error Resume Next
        rsChk.Open sqlChk, connStr   ' ← konsisten connStr
        docFound = False
        If Err.Number <> 0 Then
            dbErrDesc = Err.Description
            Err.Clear
            ' Jika DB error saat cek, anggap tidak ada → pakai nomor ini
            ' (lebih aman daripada infinite loop)
            docFound = False
        ElseIf Not rsChk.EOF Then
            If CInt(Nz(rsChk("cnt"), 0)) > 0 Then
                docFound = True
            End If
        End If
        On Error GoTo 0
        If rsChk.State = 1 Then rsChk.Close
        Set rsChk = Nothing
    Loop

    '=====================================================
    ' LANGKAH 3: Return hasil
    '=====================================================
    If docFound Then
        pageMsg = pageMsg & "<div class='msg-box msg-error'>" & _
                  "GenerateDocNo GAGAL setelah " & tryCount & " percobaan. " & _
                  "AutoNum terakhir dicoba: <b>" & AutoNum & "</b>" & _
                  "</div>"
        GenerateDocNo = ""
    Else
        GenerateDocNo = AutoNum
    End If
End Function

'=============================================================
' INISIALISASI
'=============================================================
uploadsDirVar = Server.MapPath(".") & "\tempUploads\"

Dim fsoInit
Set fsoInit = Server.CreateObject("Scripting.FileSystemObject")
If Not fsoInit.FolderExists(uploadsDirVar) Then
    On Error Resume Next
    fsoInit.CreateFolder(uploadsDirVar)
    On Error GoTo 0
End If
Set fsoInit = Nothing

fcCode       = "(NEW)"
fcfilename   = ""
namefile     = ""
txtYear      = Request("txtYear")
txtMonth     = Request("txtMonth")
txtTable     = Request("txtTable")
txtModule    = Request("txtModule")
txtSearch    = Request("txtSearch")
txtField     = Request("txtField")
temptxtJenis = Request("txtJenis")
Remarks      = Request("Remarks")
error_flag   = Request("error_flag")
error_desc   = Request("error_desc")

'=============================================================
' FUNGSI SaveFiles
'=============================================================
Function SaveFiles()
    Dim Upload, fileKey, ks, fName
    On Error Resume Next
    Set Upload = New ASPUpload
    If Err.Number <> 0 Then
        pageMsg  = "<div class='msg-box msg-error'>Gagal inisialisasi upload: " & Err.Description & "</div>"
        Err.Clear
        SaveFiles = False
        Exit Function
    End If
    Upload.Save(uploadsDirVar)
    If Err.Number <> 0 Then
        pageMsg  = "<div class='msg-box msg-error'>Gagal menyimpan file: " & Err.Description & "</div>"
        Err.Clear
        SaveFiles = False
        Exit Function
    End If
    On Error GoTo 0

    namefile = ""
    ks = Upload.UploadedFiles.Keys
    If UBound(ks) <> -1 Then
        For Each fileKey In Upload.UploadedFiles.Keys
            fName = Upload.UploadedFiles(fileKey).FileName
            fName = Mid(fName, InStrRev(fName, "\") + 1)
            fName = Mid(fName, InStrRev(fName, "/") + 1)
            namefile = fName
        Next
    End If

    Dim tmpCode
    tmpCode = Trim(Upload.Form("fcCode"))
    If tmpCode <> "" Then
        fcCode = tmpCode
    Else
        fcCode = "(NEW)"
    End If

    fcfilename    = Upload.Form("fcfilename")
    uploaddocdate = Upload.Form("uploaddocdate")
    Remarks       = Upload.Form("Remarks")

    If Trim(txtYear)      = "" Then txtYear      = Upload.Form("txtYear")
    If Trim(txtMonth)     = "" Then txtMonth     = Upload.Form("txtMonth")
    If Trim(temptxtJenis) = "" Then temptxtJenis = Upload.Form("txtJenis")

    SaveFiles = True
End Function

'=============================================================
' FUNGSI RunSQL - Baca Excel via ACE OLEDB
'=============================================================
Function RunSQL(sqlStr, rsOut)
    Dim myRs, connExcel
    connExcel = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                "Data Source=" & uploadsDirVar & namefile & ";" & _
                "Extended Properties=""Excel 12.0;IMEX=1;HDR=YES;"""
    On Error Resume Next
    Set myRs = CreateObject("ADODB.Recordset")
    myRs.Open sqlStr, connExcel, 1, 1
    If Err.Number <> 0 Then
        RunSQL  = False
        pageMsg = pageMsg & "<div class='msg-box msg-error'>Error membaca Excel: " & Err.Description & "<br>" & _
                  "Pastikan file tidak sedang dibuka dan nama sheet adalah <b>Sheet1</b>.</div>"
        Err.Clear
        Set myRs = Nothing
        Exit Function
    End If
    On Error GoTo 0
    Set rsOut = myRs
    RunSQL = True
End Function

'=============================================================
' FUNGSI OutputForm
'=============================================================
Function OutputForm()
%>
<% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-upload" aria-hidden="true"></i>Upload BA Fingerprint (Excel)</div><% end if %>
<form name="frmSend" accept-charset="utf-8" enctype="multipart/form-data"
      onsubmit="return onSubmitForm();" method="post"
      action="upload_ba_fp_02.asp?txtYear=<%=txtYear%>&txtMonth=<%=txtMonth%>">
    <table class="data" width="100%">
        <tr>
            <td align="left" width="180">No. Document</td>
            <td align="left">
                <input type="text" maxlength="15" name="fcCode" size="30"
                       value="<%=fcCode%>" readonly style="background:#999999;">
                <font color="#FF0000">* (Diisi otomatis oleh sistem)</font>
            </td>
        </tr>
        <tr>
            <td align="left">Nama File (.xls / .xlsx)</td>
            <td align="left">
                <input type="file" name="txtFile" accept=".xls,.xlsx">
                <font color="#FF0000">*</font>
            </td>
        </tr>
        <tr>
            <td colspan="2" align="left" style="padding-top:6px;">
                <input type="hidden" value="True"            name="cmdView">
                <input type="hidden" value="<%=txtTable%>"   name="txtTable">
                <input type="hidden" value="<%=txtModule%>"  name="txtModule">
                <input type="hidden" value="<%=txtMonth%>"   name="txtMonth">
                <input type="hidden" value="<%=txtYear%>"    name="txtYear">
                <input type="hidden" name="txtJenis"         value="New">
                <input type="submit" name="cmdSubmit"        value="Upload">
            </td>
        </tr>
    </table>
</form>
<%
End Function

'=============================================================
' MAIN LOGIC
'=============================================================
If Request.ServerVariables("REQUEST_METHOD") <> "POST" Then

    OutputForm()

Else

    Dim saveOK : saveOK = SaveFiles()
    OutputForm()

    If Not saveOK Then
        pageMsg     = pageMsg & "<div class='msg-box msg-error'>Proses upload gagal. Silakan coba lagi.</div>"
        stopProcess = True
    End If

    If Not stopProcess Then
        If namefile = "" Then
            pageMsg     = pageMsg & "<div class='msg-box msg-error'><b>Tidak ada file yang diterima server.</b><br>Pastikan file sudah dipilih.</div>"
            stopProcess = True
        End If
    End If

    If Not stopProcess Then
        If Not IsValidExcelFile(namefile) Then
            pageMsg     = pageMsg & "<div class='msg-box msg-error'>File <b>" & namefile & "</b> bukan format Excel (.xls / .xlsx).</div>"
            stopProcess = True
        End If
    End If

    If Not stopProcess Then
        If Trim(txtYear) = "" Or Trim(txtMonth) = "" Then
            pageMsg     = pageMsg & "<div class='msg-box msg-error'>Parameter Year/Month tidak ditemukan.</div>"
            stopProcess = True
        End If
    End If

    If Not stopProcess Then
        pageMsg = pageMsg & "<div class='msg-box msg-info'>File diterima: <b>" & namefile & "</b></div>"
    End If

    '--- Cek dokumen existing ---
    If Not stopProcess Then
        If Trim(fcCode) <> "" And Trim(fcCode) <> "(NEW)" Then
            Dim rsCekDoc, strCekDoc
            Set rsCekDoc = Server.CreateObject("ADODB.Recordset")
            strCekDoc = " SELECT docno FROM userfinger.uploadheader_mf_tab" & _
                        " WHERE docno = '" & SafeSQL(fcCode) & "'"
            On Error Resume Next
            rsCekDoc.Open strCekDoc, connstr
            If Err.Number <> 0 Then
                pageMsg     = pageMsg & "<div class='msg-box msg-error'>Error cek dokumen: " & Err.Description & "</div>"
                stopProcess = True
                Err.Clear
            ElseIf Not rsCekDoc.EOF Then
                pageMsg = pageMsg & "<script>window.location.href='upload_ba_fp_01.asp?fcCode=" & _
                          Server.URLEncode(fcCode) & "&cmdView=True';</script>"
                stopProcess = True
            End If
            On Error GoTo 0
            If Not rsCekDoc Is Nothing Then
                If rsCekDoc.State = 1 Then rsCekDoc.Close
                Set rsCekDoc = Nothing
            End If
        End If
    End If

    '=========================================================
    ' PROSES UPLOAD
    '=========================================================
    If Not stopProcess And temptxtJenis = "New" Then

        uploaddocdate = txtMonth & "/1/" & txtYear

        Dim temp_ba, temp_empno, temp_tdate, temp_timein
        Dim temp_timeout, temp_tdateout, temp_reasontype
        Dim cek_error, error_msg
        Dim empInfo, emp_fcba
        Dim totalRows, validRows, skippedRows

        sql          = "SELECT * FROM [Sheet1$]"
        cek_error    = False
        error_msg    = ""
        totalRows    = 0
        validRows    = 0
        skippedRows  = 0

        Dim dataArr()
        ReDim dataArr(0)
        Dim dataCount : dataCount = 0

        '=====================================================
        ' LANGKAH 1: BACA EXCEL & VALIDASI
        '=====================================================
        If RunSQL(sql, rs) Then

            Do While Not rs.EOF
                temp_ba         = Trim(CStr(Nz(rs.Fields.Item(0).Value, "")))
                temp_empno      = Trim(CStr(Nz(rs.Fields.Item(1).Value, "")))
                temp_tdate = Trim(CStr(Nz(rs.Fields.Item(2).Value, "")))
                If IsDate(temp_tdate) Then
                    Dim dtTemp1 : dtTemp1 = CDate(temp_tdate)
                    temp_tdate = Right("0" & Month(dtTemp1), 2) & "/" & _
                                Right("0" & Day(dtTemp1), 2)   & "/" & _
                                Year(dtTemp1)
                ElseIf IsNumeric(temp_tdate) Then
                    Dim dtTemp2 : dtTemp2 = CDate(CDbl(temp_tdate))
                    temp_tdate = Right("0" & Month(dtTemp2), 2) & "/" & _
                                Right("0" & Day(dtTemp2), 2)   & "/" & _
                                Year(dtTemp2)
                End If
                temp_timein     = Trim(CStr(Nz(rs.Fields.Item(3).Value, "")))
                temp_timeout    = Trim(CStr(Nz(rs.Fields.Item(4).Value, "")))
                temp_reasontype = Trim(CStr(Nz(rs.Fields.Item(5).Value, "")))
                temp_tdateout   = Trim(CStr(Nz(rs.Fields.Item(6).Value, "")))

                ' Skip baris kosong sepenuhnya
                If temp_ba = "" And temp_empno = "" Then
                    rs.MoveNext
                Else
                    totalRows = totalRows + 1
                    Dim rowOK : rowOK = True

                    ' Validasi kolom wajib
                    If temp_ba = "" Then
                        error_msg = error_msg & "Baris " & totalRows & ": NOBA kosong.<br>"
                        rowOK = False : cek_error = True
                    End If
                    If temp_empno = "" Then
                        error_msg = error_msg & "Baris " & totalRows & ": EMPNO kosong.<br>"
                        rowOK = False : cek_error = True
                    End If
                    If temp_tdate = "" Then
                        error_msg = error_msg & "Baris " & totalRows & ": tdate kosong.<br>"
                        rowOK = False : cek_error = True
                    End If
                    If temp_timein = "" Then
                        error_msg = error_msg & "Baris " & totalRows & ": timein kosong.<br>"
                        rowOK = False : cek_error = True
                    End If
                    If temp_timeout = "" Then
                        error_msg = error_msg & "Baris " & totalRows & ": timeout kosong.<br>"
                        rowOK = False : cek_error = True
                    End If

                    ' Validasi ba_reasontype
                    If temp_reasontype <> "1" And temp_reasontype <> "2" And _
                       temp_reasontype <> "3" And temp_reasontype <> "4" And _
                       temp_reasontype <> "5" And temp_reasontype <> "6" Then
                        error_msg = error_msg & "Baris " & totalRows & ": ba_reasontype '" & _
                                    temp_reasontype & "' tidak valid (harus 1-6).<br>"
                        rowOK = False : cek_error = True
                    End If

                    ' Validasi periode: tdate harus sesuai bulan & tahun yang dipilih.
                    ' Baris di luar periode di-SKIP (tidak membatalkan upload).
                    If rowOK And temp_tdate <> "" Then
                        On Error Resume Next
                        Dim dtCheck : dtCheck = CDate(temp_tdate)
                        If Err.Number <> 0 Then
                            error_msg = error_msg & "<b>Baris " & totalRows & "</b>: format tdate '" & _
                                        temp_tdate & "' tidak valid &mdash; dilewati.<br>"
                            rowOK = False
                            Err.Clear
                        End If
                        On Error GoTo 0
                    End If

                    If rowOK Then
                        ' Ambil fcba dari ATTENDANCE_TAB, fallback ke session
                        empInfo  = GetEmpInfo(temp_empno)
                        emp_fcba = empInfo(0)
                        If emp_fcba = "" Then emp_fcba = Session("fcBA")

                        ReDim Preserve dataArr(dataCount)
                        dataArr(dataCount) = temp_ba         & "|~|" & _
                                             temp_empno      & "|~|" & _
                                             temp_tdate      & "|~|" & _
                                             temp_timein     & "|~|" & _
                                             temp_timeout    & "|~|" & _
                                             temp_reasontype & "|~|" & _
                                             temp_tdateout   & "|~|" & _
                                             emp_fcba
                        dataCount = dataCount + 1
                        validRows = validRows + 1
                    Else
                        skippedRows = skippedRows + 1
                    End If

                    rs.MoveNext
                End If
            Loop
            rs.Close
            Set rs = Nothing

        Else
            stopProcess = True
        End If

        ' Tampilkan ringkasan baca Excel
        If Not stopProcess Then
            Dim skipInfo : skipInfo = ""
            If skippedRows > 0 Then
                skipInfo = " | Dilewati: <b>" & skippedRows & "</b>"
            End If
            pageMsg = pageMsg & "<div class='msg-box msg-info'>Total baris terbaca: <b>" & _
                      totalRows & "</b> | Valid: <b>" & validRows & "</b>" & skipInfo & "</div>"
            If skippedRows > 0 And error_msg <> "" Then
                pageMsg   = pageMsg & "<div class='msg-box msg-warning'><b>Detail baris dilewati:</b><br>" & error_msg & "</div>"
                error_msg = ""
            End If
        End If

        ' Batalkan jika ada error kolom wajib
        If Not stopProcess And cek_error Then
            pageMsg     = pageMsg & "<div class='msg-box msg-error'><b>Upload dibatalkan:</b><br>" & error_msg & "</div>"
            stopProcess = True
        End If

        If Not stopProcess And validRows = 0 Then
            pageMsg     = pageMsg & "<div class='msg-box msg-error'>Tidak ada data valid. " & _
                          "Pastikan file tidak kosong dan sheet bernama <b>Sheet1</b>.</div>"
            stopProcess = True
        End If

        '=====================================================
        ' LANGKAH 2: GENERATE DOCNO UNIK
        ' Ambil MAX angka di ujung DOCNO dari tabel untuk
        ' kombinasi fcba + tahun + bulan, lalu increment +1.
        ' Loop cek ke tabel sampai dapat yang belum ada.
        '=====================================================
        If Not stopProcess Then
            AutoNum = GenerateDocNo(Session("fcBA"), txtYear, txtMonth)
            If AutoNum = "" Then
                pageMsg     = pageMsg & "<div class='msg-box msg-error'>Gagal mendapatkan nomor dokumen unik. " & _
                              "Hubungi administrator.</div>"
                stopProcess = True
            Else
                fcCode  = AutoNum
                pageMsg = pageMsg & "<div class='msg-box msg-info'>Nomor dokumen: <b>" & fcCode & "</b></div>"
            End If
        End If

        '=====================================================
        ' LANGKAH 3: INSERT HEADER
        '=====================================================
        If Not stopProcess Then
            q = " INSERT INTO userfinger.uploadheader_mf_tab("    & _
                "   docno, tdate, filename, fcentry, fcip, fcba," & _
                "   rowstateid, lastupdate, lasttime)"             & _
                " VALUES ("                                        & _
                "   '" & SafeSQL(fcCode)   & "',"                 & _
                "   TO_DATE('" & SafeSQL(uploaddocdate) & "','mm/dd/yyyy')," & _
                "   '" & SafeSQL(namefile) & "',"                 & _
                "   '" & SafeSQL(Session("CodeName")) & "',"       & _
                "   '" & SafeSQL(Request.ServerVariables("REMOTE_ADDR")) & "'," & _
                "   '" & SafeSQL(Session("fcBA")) & "',"           & _
                "   'Planned',"                                    & _
                "   SYSDATE,"                                      & _
                "   TO_CHAR(SYSDATE,'hh24:mi'))"
            On Error Resume Next
            connStr.Execute(q)
            If Err.Number <> 0 Then
                pageMsg     = pageMsg & "<div class='msg-box msg-error'>Gagal insert header: " & Err.Description & "</div>"
                stopProcess = True
                Err.Clear
            End If
            On Error GoTo 0
        End If

        '=====================================================
        ' LANGKAH 4: MERGE DETAIL PER BARIS
        ' Key MERGE: empno + tdate + fcba
        '   MATCHED     → UPDATE (tidak duplikasi)
        '   NOT MATCHED → INSERT baru
        '=====================================================
        If Not stopProcess Then
            Dim processed, errCount
            Dim sql_timein, sql_timeout, sql_tdate_out
            Dim arrRow, rowFields, rowStatus
            processed = 0
            errCount  = 0

            Dim tblDetail
            tblDetail = "<table class='tbl-result'>" & _
                        "<tr><th>No</th><th>NOBA</th><th>EMPNO</th><th>Tanggal</th>" & _
                        "<th>Time In</th><th>Time Out</th><th>FCBA</th><th>Status</th></tr>"

            Dim rowNum : rowNum = 0
            For arrRow = 0 To dataCount - 1
                rowNum    = rowNum + 1
                rowFields = Split(dataArr(arrRow), "|~|")

                If UBound(rowFields) >= 7 Then
                    temp_ba         = rowFields(0)
                    temp_empno      = rowFields(1)
                    temp_tdate      = rowFields(2)
                    temp_timein     = rowFields(3)
                    temp_timeout    = rowFields(4)
                    temp_reasontype = rowFields(5)
                    temp_tdateout   = rowFields(6)
                    emp_fcba        = rowFields(7)

                    sql_tdate_out = temp_tdateout
                    If Trim(sql_tdate_out) = "" Then sql_tdate_out = temp_tdate

                    sql_timein  = "TO_DATE('" & SafeSQL(temp_tdate)    & " " & SafeSQL(temp_timein)  & ":00','mm/dd/yyyy hh24:mi:ss')"
                    sql_timeout = "TO_DATE('" & SafeSQL(sql_tdate_out) & " " & SafeSQL(temp_timeout) & ":00','mm/dd/yyyy hh24:mi:ss')"

                    q = " MERGE INTO userfinger.uploaddetail_mf_tab e"                & _
                        " USING ("                                                    & _
                        "   SELECT"                                                   & _
                        "     '" & SafeSQL(fcCode)           & "' AS docno,"          & _
                        "     '" & SafeSQL(temp_ba)          & "' AS no_ba,"          & _
                        "     '" & SafeSQL(temp_empno)       & "' AS empno,"          & _
                        "     TO_DATE('" & SafeSQL(temp_tdate) & "','mm/dd/yyyy') AS tdate," & _
                        "     " & sql_timein                 & " AS timein,"          & _
                        "     " & sql_timeout                & " AS timeout,"         & _
                        "     '" & SafeSQL(temp_reasontype)  & "' AS ba_reasontype,"  & _
                        "     '" & SafeSQL(emp_fcba)         & "' AS fcba,"           & _
                        "     SYSDATE                           AS lastupdate,"       & _
                        "     TO_CHAR(SYSDATE,'hh24:mi')        AS lasttime,"         & _
                        "     'FALSE'                           AS error_flag,"       & _
                        "     NULL                              AS error_desc"        & _
                        "   FROM DUAL"                                                & _
                        " ) h"                                                        & _
                        " ON (e.docno = h.docno AND e.empno = h.empno AND e.tdate = h.tdate AND e.fcba = h.fcba)" & _
                        " WHEN MATCHED THEN UPDATE SET"                               & _
                        "   e.no_ba         = h.no_ba,"                               & _
                        "   e.timein        = h.timein,"                              & _
                        "   e.timeout       = h.timeout,"                             & _
                        "   e.ba_reasontype = h.ba_reasontype,"                       & _
                        "   e.error_flag    = h.error_flag,"                          & _
                        "   e.error_desc    = h.error_desc,"                          & _
                        "   e.lastupdate    = h.lastupdate,"                          & _
                        "   e.lasttime      = h.lasttime"                             & _
                        " WHEN NOT MATCHED THEN INSERT ("                             & _
                        "   docno, no_ba, empno, tdate, timein, timeout,"             & _
                        "   ba_reasontype, fcba,"                                     & _
                        "   lastupdate, lasttime, error_flag, error_desc"             & _
                        " ) VALUES ("                                                 & _
                        "   h.docno, h.no_ba, h.empno, h.tdate, h.timein, h.timeout," & _
                        "   h.ba_reasontype, h.fcba,"                                 & _
                        "   h.lastupdate, h.lasttime, h.error_flag, h.error_desc"     & _
                        " )"

                    On Error Resume Next
                    connStr.Execute(q)
                    If Err.Number <> 0 Then
                        rowStatus = "<font color='red'>GAGAL: " & Err.Description & "</font>"
                        errCount  = errCount + 1
                        Err.Clear
                    Else
                        rowStatus = "<font color='green'>&#10003; OK</font>"
                        processed = processed + 1
                    End If
                    On Error GoTo 0

                    tblDetail = tblDetail & "<tr>"          & _
                                "<td>" & rowNum        & "</td>" & _
                                "<td>" & temp_ba       & "</td>" & _
                                "<td>" & temp_empno    & "</td>" & _
                                "<td>" & temp_tdate    & "</td>" & _
                                "<td>" & temp_timein   & "</td>" & _
                                "<td>" & temp_timeout  & "</td>" & _
                                "<td>" & emp_fcba      & "</td>" & _
                                "<td>" & rowStatus     & "</td>" & _
                                "</tr>"
                End If
            Next
            tblDetail = tblDetail & "</table>"

            ' Ringkasan hasil
            If errCount = 0 And processed > 0 Then
                pageMsg = pageMsg & "<div class='msg-box msg-success'>" & _
                          "&#10003; <b>Import BERHASIL!</b> &nbsp;|&nbsp; " & _
                          "Dokumen: <b>" & fcCode & "</b> &nbsp;|&nbsp; " & _
                          "Total baris: <b>" & processed & "</b></div>"
            ElseIf processed > 0 And errCount > 0 Then
                pageMsg = pageMsg & "<div class='msg-box msg-warning'>" & _
                          "&#9888; <b>Import SEBAGIAN BERHASIL</b> &nbsp;|&nbsp; " & _
                          "Berhasil: <b>" & processed & "</b> &nbsp;|&nbsp; " & _
                          "Gagal: <b>" & errCount & "</b></div>"
            Else
                pageMsg = pageMsg & "<div class='msg-box msg-error'>" & _
                          "&#10007; <b>Import GAGAL!</b> Tidak ada baris yang berhasil.</div>"
            End If

            pageMsg = pageMsg & tblDetail

            ' Redirect via JS agar session tetap hidup
            ' If processed > 0 Then
            '     pageMsg = pageMsg & "<div class='msg-box msg-info'>" & _
            '               "Mengalihkan ke halaman detail dalam 3 detik... " & _
            '               "<a href='upload_ba_fp_01.asp?fcCode=" & fcCode & "&cmdView=True'>Klik di sini</a> jika tidak otomatis.</div>" & _
            '               "<script>setTimeout(function(){ window.location.href='upload_ba_fp_01.asp?fcCode=" & _
            '               Server.URLEncode(fcCode) & "&cmdView=True'; }, 3000);</script>"
            ' End If
        End If

    ElseIf Not stopProcess Then
        cmdView = Request("cmdView")
    End If

    Response.Write pageMsg

End If
%>
    </div>
</div>
<!--#include file="bottom.jsp"-->
</body>
</html>
