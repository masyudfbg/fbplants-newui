<%
Option Explicit
Response.Buffer = True
%>
<!--#include file="connDb1.inc"-->
<%
' ===========================================================
' simpan_upload_stockopname.asp
' - Full multipart parser (no Persits)
' - Field file name: filestockopname
' - Preview, Validasi, Import, Auto-delete per STORE|YEAR|MONTH
' ===========================================================

' -----------------------
' Konstanta
' -----------------------
Const adTypeBinary = 1
Const adTypeText = 2
Const adSaveCreateOverWrite = 2

' -----------------------
' Helper functions
' -----------------------
Function SafeTrim(val)
    If IsNull(val) Then
        SafeTrim = ""
    Else
        SafeTrim = Trim(CStr(val))
    End If
End Function

Function EscapeSQL(val)
    If IsNull(val) Then
        EscapeSQL = ""
    Else
        EscapeSQL = Replace(CStr(val), "'", "''")
    End If
End Function

Function ToDoubleOrZero(val)
    If IsNumeric(val) Then
        ToDoubleOrZero = CDbl(val)
    Else
        ToDoubleOrZero = 0
    End If
End Function

' -----------------------
' SaveBinaryToFile(binary, path)
' -----------------------
Sub SaveBinaryToFile(binaryData, filePath)
    Dim stream
    Set stream = Server.CreateObject("ADODB.Stream")
    stream.Type = adTypeBinary
    stream.Open
    stream.Write binaryData
    stream.SaveToFile filePath, adSaveCreateOverWrite
    stream.Close
    Set stream = Nothing
End Sub

' -----------------------
' ReadFileAsText(path) with iso-8859-1
' -----------------------
Function ReadFileAsTextIso(path)
    Dim stm
    Set stm = Server.CreateObject("ADODB.Stream")
    stm.Type = adTypeText
    stm.Charset = "iso-8859-1"
    stm.Open
    stm.LoadFromFile path
    ReadFileAsTextIso = stm.ReadText
    stm.Close
    Set stm = Nothing
End Function

' -----------------------
' SaveTextAsBinaryFileIso(text, path)
' writes text (iso-8859-1) so bytes preserved
' -----------------------
Sub SaveTextAsBinaryFileIso(txt, path)
    Dim stm
    Set stm = Server.CreateObject("ADODB.Stream")
    stm.Type = adTypeText
    stm.Charset = "iso-8859-1"
    stm.Open
    stm.WriteText txt
    stm.SaveToFile path, adSaveCreateOverWrite
    stm.Close
    Set stm = Nothing
End Sub

Function Nz(v)
    If IsNull(v) Or v = "" Then
        Nz = ""
    Else
        Nz = v
    End If
End Function

' -----------------------
' Extract file part from raw multipart text
' returns "" if not found, otherwise returns dictionary with keys: filename, content (text preserving bytes)
' -----------------------
Function ExtractFilePart(rawText, fieldName)
    Dim result
    Set result = Nothing

    Dim parts, i, p, headerSection, fileName, posHeaderEnd, contentPart

    ' Normalize line endings to vbCrLf
    rawText = Replace(rawText, vbCr, vbCr)
    ' boundary is first line until CRLF
    Dim firstLineEnd
    firstLineEnd = InStr(rawText, vbCrLf)
    If firstLineEnd = 0 Then
        ExtractFilePart = ""
        Exit Function
    End If
    Dim boundary
    boundary = Mid(rawText, 1, firstLineEnd - 1)

    ' split parts by boundary
    parts = Split(rawText, boundary)

    For i = 0 To UBound(parts)
        p = parts(i)
        If InStr(p, "name=""" & fieldName & """") > 0 Then
            ' find filename
            fileName = ""
            headerSection = ""
            posHeaderEnd = InStr(p, vbCrLf & vbCrLf)
            If posHeaderEnd > 0 Then
                headerSection = Left(p, posHeaderEnd - 1)
                ' cari filename="
                If InStr(headerSection, "filename=""") > 0 Then
                    Dim fstart, fend
                    fstart = InStr(headerSection, "filename=""") + Len("filename=""")
                    fend = InStr(fstart, headerSection, """")
                    If fend > fstart Then
                        fileName = Mid(headerSection, fstart, fend - fstart)
                    End If
                End If

                ' content adalah sisanya sampai akhir part (hilangkan trailing CRLF)
                contentPart = Mid(p, posHeaderEnd + 4) ' setelah double CRLF
                ' remove trailing CRLF jika ada
                If Right(contentPart, 2) = vbCrLf Then
                    contentPart = Left(contentPart, Len(contentPart) - 2)
                End If

                ' set result dictionary
                Dim dict
                Set dict = Server.CreateObject("Scripting.Dictionary")
                dict.Add "filename", fileName
                dict.Add "content", contentPart
                Set ExtractFilePart = dict
                Exit Function
            End If
        End If
    Next

    ExtractFilePart = ""
End Function

' -----------------------
' Main flow
' -----------------------
Dim mode
mode = Trim(Request("mode"))

' ensure temp_upload folder exists
Dim fso, tempFolderPath
Set fso = Server.CreateObject("Scripting.FileSystemObject")
tempFolderPath = Server.MapPath("temp_upload")
If Not fso.FolderExists(tempFolderPath) Then
    On Error Resume Next
    fso.CreateFolder(tempFolderPath)
    On Error GoTo 0
End If

If mode = "" Then
    ' ---------------------------
    ' UPLOAD & PREVIEW
    ' ---------------------------
    Dim totalBytes
    totalBytes = Request.TotalBytes
    If totalBytes = 0 Then
        Response.Write "<script>alert('Tidak ada data upload. Pastikan form enctype multipart/form-data dan ada file dipilih.');history.back();</script>"
        Response.End
    End If

    Dim bin
    bin = Request.BinaryRead(totalBytes)

    ' save raw request to temp file
    Dim reqTmpFile
    reqTmpFile = tempFolderPath & "\req_" & Replace(Replace(Replace(CStr(Now()), ":", ""), " ", "_"), "/", "") & ".tmp"
    SaveBinaryToFile bin, reqTmpFile

    ' read as text iso-8859-1
    Dim rawText
    rawText = ReadFileAsTextIso(reqTmpFile)

    ' extract file part for field 'filestockopname'
    Dim filePart
    Set filePart = ExtractFilePart(rawText, "filestockopname")
    If IsEmpty(filePart) Or filePart = "" Then
        Response.Write "<h3 style='color:red'>File tidak ditemukan pada field 'filestockopname'. Pastikan input file punya name='filestockopname'.</h3>"
        Response.End
    End If

    Dim origFileName
    origFileName = filePart("filename")
    If origFileName = "" Then
        Response.Write "<h3 style='color:red'>Nama file tidak terdeteksi dalam upload.</h3>"
        Response.End
    End If

    ' validate extension .xlsx
    If LCase(Right(origFileName, 5)) <> ".xlsx" Then
        Response.Write "<script>alert('File harus berekstensi .xlsx');history.back();</script>"
        Response.End
    End If

    ' rename file
    Dim timestamp, savedFileName, savedPath
    timestamp = Replace(Replace(Replace(Replace(CStr(Now()), ":", ""), " ", "_"), "/", ""), "-", "")
    savedFileName = "UPLOAD_STOCKOPNAME_" & timestamp & ".xlsx"
    savedPath = tempFolderPath & "\" & savedFileName

    ' save content (content is text preserving bytes) as binary xlsx
    SaveTextAsBinaryFileIso filePart("content"), savedPath

    ' remove request tmp
    On Error Resume Next
    If fso.FileExists(reqTmpFile) Then fso.DeleteFile reqTmpFile, True
    On Error GoTo 0

    ' now read excel with OLEDB for preview
    Dim strExcelConn, excelConn, rsExcel
    strExcelConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & savedPath & ";Extended Properties=""Excel 12.0 Xml;HDR=YES"";"

    On Error Resume Next
    Set excelConn = Server.CreateObject("ADODB.Connection")
    excelConn.Open strExcelConn
    If Err.Number <> 0 Then
        Response.Write "<h3 style='color:red'>Error membuka file Excel. Pastikan Provider Microsoft.ACE.OLEDB.12.0 tersedia di server.</h3>"
        Response.Write "<p>Detail: " & Server.HTMLEncode(Err.Description) & "</p>"
        Response.End
    End If
    On Error GoTo 0

    On Error Resume Next
    Set rsExcel = excelConn.Execute("SELECT * FROM [Sheet1$]")
    If Err.Number <> 0 Then
        Response.Write "<h3 style='color:red'>Error membaca Sheet1. Pastikan sheet bernama 'Sheet1' dan kolom sesuai template.</h3>"
        Response.End
    End If
    On Error GoTo 0

    ' Render preview
    Response.Write "<h2>Preview Data Stock Opname</h2>"
    Response.Write "<p>File: " & Server.HTMLEncode(savedFileName) & "</p>"
    Response.Write "<form method='post' action='simpan_upload_stockopname.asp'>"
    Response.Write "<input type='hidden' name='mode' value='import'>"
    Response.Write "<input type='hidden' name='filename' value='" & Server.HTMLEncode(savedFileName) & "'>"

    Response.Write "<table border='1' cellpadding='4' cellspacing='0' style='border-collapse:collapse;'>"
    Response.Write "<tr style='background:#dddddd'>"
    Response.Write "<th>NO</th><th>STORECODE</th><th>ITEMCODE</th><th>YEAR</th><th>MONTH</th><th>QTY</th><th>AMTPERUNIT</th><th>FCBA</th><th>AMOUNT</th><th>STATUS</th>"
    Response.Write "</tr>"

    Dim rowNo
    rowNo = 0

    Do Until rsExcel.EOF
        rowNo = rowNo + 1
        Dim sc, ic, yr, mon, q, apu, fcb, amt
        sc = SafeTrim(rsExcel("STORECODE"))
        ic = SafeTrim(rsExcel("ITEMCODE"))
        yr = SafeTrim(rsExcel("YEAR"))
        mon = SafeTrim(rsExcel("MONTH"))
        q = SafeTrim(rsExcel("QTY"))
        apu = SafeTrim(rsExcel("AMTPERUNIT"))
        fcb = SafeTrim(rsExcel("FCBA"))
        amt = SafeTrim(rsExcel("AMOUNT"))

        Dim em
        em = ""
        If sc = "" Then em = em & "STORECODE kosong; "
        If ic = "" Then em = em & "ITEMCODE kosong; "
        If yr = "" Then em = em & "YEAR kosong; "
        If mon = "" Then em = em & "MONTH kosong; "
        If q = "" Or Not IsNumeric(q) Then em = em & "QTY bukan angka; "
        If apu <> "" And Not IsNumeric(apu) Then em = em & "AMTPERUNIT bukan angka; "
        If amt <> "" And Not IsNumeric(amt) Then em = em & "AMOUNT bukan angka; "

        If em <> "" Then
            Response.Write "<tr style='background:#ffdddd'>"
        Else
            Response.Write "<tr>"
        End If

        Response.Write "<td align='center'>" & rowNo & "</td>"
        Response.Write "<td>" & Server.HTMLEncode(sc) & "</td>"
        Response.Write "<td>" & Server.HTMLEncode(ic) & "</td>"
        Response.Write "<td align='center'>" & Server.HTMLEncode(yr) & "</td>"
        Response.Write "<td align='center'>" & Server.HTMLEncode(mon) & "</td>"
        Response.Write "<td align='right'>" & Server.HTMLEncode(q) & "</td>"
        Response.Write "<td align='right'>" & Server.HTMLEncode(apu) & "</td>"
        Response.Write "<td>" & Server.HTMLEncode(fcb) & "</td>"
        Response.Write "<td align='right'>" & Server.HTMLEncode(amt) & "</td>"

        If em <> "" Then
            Response.Write "<td style='color:#a10000'>" & Server.HTMLEncode(em) & "</td>"
        Else
            Response.Write "<td style='color:green'>VALID</td>"
        End If

        Response.Write "</tr>"

        rsExcel.MoveNext
    Loop

    If rowNo = 0 Then
        Response.Write "<tr><td colspan='10' style='color:red'>Tidak ditemukan data di Sheet1.</td></tr>"
    End If

    Response.Write "</table><br><br>"
    Response.Write "<input type='submit' value='IMPORT SEKARANG' style='padding:6px 12px'>"
    Response.Write " &nbsp; <a href='javascript:history.back()'>Batal / Kembali</a>"
    Response.Write "</form>"

    rsExcel.Close
    excelConn.Close
    Set rsExcel = Nothing
    Set excelConn = Nothing

    Response.End
End If

' ---------------------------
' MODE = import -> lakukan insert
' ---------------------------
If LCase(mode) = "import" Then
    Dim fname
    fname = Trim(Request("filename"))
    If fname = "" Then
        Response.Write "<script>alert('Filename tidak ditemukan. Lakukan upload ulang.');history.back();</script>"
        Response.End
    End If

    Dim excelFullPath
    excelFullPath = tempFolderPath & "\" & fname

    ' buka excel
    Dim strExcelConn2, excelConn2, rsExcel2
    strExcelConn2 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & excelFullPath & ";Extended Properties=""Excel 12.0 Xml;HDR=YES"";"
    On Error Resume Next
    Set excelConn2 = Server.CreateObject("ADODB.Connection")
    excelConn2.Open strExcelConn2
    If Err.Number <> 0 Then
        Response.Write "<h3 style='color:red'>Error membuka file Excel saat import: " & Server.HTMLEncode(Err.Description) & "</h3>"
        Response.End
    End If
    On Error GoTo 0

    Set rsExcel2 = excelConn2.Execute("SELECT * FROM [Sheet1$]")

    ' buka koneksi DB
    Dim connDB
    Set connDB = Server.CreateObject("ADODB.Connection")
    connDB.Open strConnDb1

    ' kumpulkan kombinasi untuk delete
    Dim dictKeys
    Set dictKeys = Server.CreateObject("Scripting.Dictionary")

    Dim total, sukses, gagal
    total = 0 : sukses = 0 : gagal = 0

    ' baca file ke memori (array) untuk mengetahui kombinasi delete dulu
    Dim rowsData()
    ReDim rowsData(0)
    Dim idx
    idx = 0

    Do Until rsExcel2.EOF
        total = total + 1

        Dim r_store, r_item, r_year, r_month, r_qty, r_amtper, r_fcba, r_amount
        r_store  = SafeTrim(Nz(rsExcel2("STORECODE")))
        r_item   = SafeTrim(Nz(rsExcel2("ITEMCODE")))
        r_year   = SafeTrim(Nz(rsExcel2("YEAR")))
        r_month  = SafeTrim(Nz(rsExcel2("MONTH")))
        r_qty    = SafeTrim(Nz(rsExcel2("QTY")))
        r_amtper = SafeTrim(Nz(rsExcel2("AMTPERUNIT")))
        r_fcba   = SafeTrim(Nz(rsExcel2("FCBA")))
        r_amount = SafeTrim(Nz(rsExcel2("AMOUNT")))

        If idx = 0 Then ReDim rowsData(8, 1)

        idx = idx + 1
        ReDim Preserve rowsData(8, idx)
        rowsData(1, idx) = r_store
        rowsData(2, idx) = r_item
        rowsData(3, idx) = r_year
        rowsData(4, idx) = r_month
        rowsData(5, idx) = r_qty
        rowsData(6, idx) = r_amtper
        rowsData(7, idx) = r_fcba
        rowsData(8, idx) = r_amount

        If r_store <> "" And r_year <> "" And r_month <> "" Then
            Dim key
            key = r_store & "|" & r_year & "|" & r_month
            If Not dictKeys.Exists(key) Then dictKeys.Add key, 1
        End If

        rsExcel2.MoveNext
    Loop

    ' delete per kombinasi
    Dim eachKey
    For Each eachKey In dictKeys.Keys
        Dim parts, ks, ky, km
        parts = Split(eachKey, "|")
        ks = Replace(parts(0), "'", "''")
        ky = Replace(parts(1), "'", "''")
        km = Replace(parts(2), "'", "''")
        Dim delSql
        delSql = "DELETE FROM upload_stockopname WHERE STORECODE='" & ks & "' AND YEAR='" & ky & "' AND MONTH='" & km & "'"
        On Error Resume Next
        connDB.Execute delSql
        On Error GoTo 0
    Next

    ' insert valid rows
    Dim i
    For i = 1 To idx
        Dim vs, vi, vy, vm, vq, vapu, vfcba, vamt
        vs = rowsData(1, i)
        vi = rowsData(2, i)
        vy = rowsData(3, i)
        vm = rowsData(4, i)
        vq = rowsData(5, i)
        vapu = rowsData(6, i)
        vfcba = rowsData(7, i)
        vamt = rowsData(8, i)

        Dim validRow
        validRow = True
        If vs = "" Then validRow = False
        If vi = "" Then validRow = False
        If vy = "" Then validRow = False
        If vm = "" Then validRow = False
        If vq = "" Or Not IsNumeric(vq) Then validRow = False

        If validRow Then
            Dim insSql
            insSql = "INSERT INTO upload_stockopname (STORECODE, ITEMCODE, YEAR, MONTH, QTY, AMTPERUNIT, FCBA, AMOUNT, FCENTRY, LASTUPDATE, LASTTIME) VALUES (" & _
                     "'" & EscapeSQL(vs) & "'," & _
                     "'" & EscapeSQL(vi) & "'," & _
                     "'" & EscapeSQL(vy) & "'," & _
                     "'" & EscapeSQL(vm) & "'," & _
                     ToDoubleOrZero(vq) & "," & _
                     ToDoubleOrZero(vapu) & "," & _
                     "'" & EscapeSQL(vfcba) & "'," & _
                     ToDoubleOrZero(vamt) & "," & _
                     "'SYSTEM', DATE(), TIME())"
            On Error Resume Next
            connDB.Execute insSql
            If Err.Number <> 0 Then
                gagal = gagal + 1
                Err.Clear
            Else
                sukses = sukses + 1
            End If
            On Error GoTo 0
        Else
            gagal = gagal + 1
        End If
    Next

    ' close connections
    rsExcel2.Close
    excelConn2.Close
    connDB.Close

    Set rsExcel2 = Nothing
    Set excelConn2 = Nothing
    Set connDB = Nothing
    Set dictKeys = Nothing

    ' delete temp xlsx
    On Error Resume Next
    If fso.FileExists(excelFullPath) Then fso.DeleteFile excelFullPath, True
    On Error GoTo 0

    ' show results
    Response.Write "<h2>HASIL IMPORT</h2>"
    Response.Write "<p>Total baris dibaca : " & total & "</p>"
    Response.Write "<p>Sukses di-insert     : <b>" & sukses & "</b></p>"
    Response.Write "<p>Gagal (skip/error)   : <b style='color:red'>" & gagal & "</b></p>"
    Response.Write "<p><a href='upload_stockopname.asp'>Kembali ke halaman upload</a></p>"
    Response.End
End If

' default
Response.Write "<h3>Mode tidak dikenal.</h3>"
%>
