<!--#include file="connDb1.Inc"-->
<%
Dim v_fcCode, v_fcba, v_date
Dim tglsplit, bulansplit, thnsplit
Dim delGLHdr, delGLDtl, delJurnalHdr, delJurnalDtl, txtSQL, qProcessPayroll
Dim raGLHdr, raGLDtl, raJurnalHdr, raJurnalDtl, raUpdate
Dim errMsg, hasError

v_fcCode = Trim(Request("fcCode"))
v_fcba   = Trim(Request("fcba"))
v_date   = Trim(Request("tgl"))

' Validasi input
If Len(v_fcCode) = 0 Or Len(v_fcba) = 0 Or Len(v_date) = 0 Then
    Response.Write "<script>alert('Parameter tidak valid!'); window.close();</script>"
    Response.End
End If

' Validasi format tanggal dd/mm/yyyy
If Not (Len(v_date) = 10 And Mid(v_date,3,1) = "/" And Mid(v_date,6,1) = "/") Then
    Response.Write "<script>alert('Format tanggal tidak valid (harus dd/mm/yyyy)!'); window.close();</script>"
    Response.End
End If

' FIX: sebelumnya bulan & hari tertukar
tglsplit    = Left(v_date, 2)       ' hari
bulansplit  = Mid(v_date, 4, 2)     ' bulan
thnsplit    = Right(v_date, 4)      ' tahun

errMsg   = ""
hasError = False

On Error Resume Next

' --- 1. Delete GLEDGERHDR ---
Err.Clear
delGLHdr = "delete from GLEDGERHDR WHERE DOCNO IN (SELECT DOCNO FROM GLEDGERDTL WHERE JOURNALNO = '" & v_fcCode & "')"
connStr.Execute delGLHdr, raGLHdr
If Err.Number <> 0 Then
    hasError = True
    errMsg = errMsg & "GLEDGERHDR: " & Err.Description & " | "
End If

' --- 2. Delete GLEDGERDTL ---
Err.Clear
delGLDtl = "delete from GLEDGERDTL where JOURNALNO = '" & v_fcCode & "'"
connStr.Execute delGLDtl, raGLDtl
If Err.Number <> 0 Then
    hasError = True
    errMsg = errMsg & "GLEDGERDTL: " & Err.Description & " | "
End If

' --- 3. Delete BKMJOURNALHDR ---
Err.Clear
delJurnalHdr = "delete from BKMJOURNALHDR where fcba = '" & v_fcba & "' and docno in ('" & v_fcCode & "','" & v_fcCode & "-INT')"
connStr.Execute delJurnalHdr, raJurnalHdr
If Err.Number <> 0 Then
    hasError = True
    errMsg = errMsg & "BKMJOURNALHDR: " & Err.Description & " | "
End If

' --- 4. Delete BKMJOURNALDTL ---
Err.Clear
delJurnalDtl = "delete from BKMJOURNALDTL where fcba = '" & v_fcba & "' and docno in ('" & v_fcCode & "','" & v_fcCode & "-INT')"
connStr.Execute delJurnalDtl, raJurnalDtl
If Err.Number <> 0 Then
    hasError = True
    errMsg = errMsg & "BKMJOURNALDTL: " & Err.Description & " | "
End If

' --- 5. Update invoiceservicehdr ---
Err.Clear
txtSQL = "UPDATE invoiceservicehdr SET PostingFlg = 0 WHERE invoiceno = '" & v_fcCode & "' AND fcba = '" & v_fcba & "'"
connStr.Execute txtSQL, raUpdate
If Err.Number <> 0 Then
    hasError = True
    errMsg = errMsg & "invoiceservicehdr: " & Err.Description & " | "
End If

' --- 6. Proses closing month ---
Err.Clear
qProcessPayroll = "BEGIN fac_closing_accounting_api.ProsesClosingMonth(" & thnsplit & "," & bulansplit & ",'" & v_fcba & "','" & Session("CodeName") & "'); END;"
connStr.Execute qProcessPayroll
If Err.Number <> 0 Then
    hasError = True
    errMsg = errMsg & "ProsesClosingMonth: " & Err.Description & " | "
End If

On Error GoTo 0

' --- Trace hasil (hapus/comment blok ini setelah selesai debug) ---
Response.Write "<!-- DEBUG TRACE" & vbCrLf
Response.Write "fcCode=" & v_fcCode & " fcba=" & v_fcba & " tgl=" & v_date & vbCrLf
Response.Write "tglsplit=" & tglsplit & " bulansplit=" & bulansplit & " thnsplit=" & thnsplit & vbCrLf
Response.Write "GLEDGERHDR affected: " & raGLHdr & vbCrLf
Response.Write "GLEDGERDTL affected: " & raGLDtl & vbCrLf
Response.Write "BKMJOURNALHDR affected: " & raJurnalHdr & " | SQL: " & delJurnalHdr & vbCrLf
Response.Write "BKMJOURNALDTL affected: " & raJurnalDtl & " | SQL: " & delJurnalDtl & vbCrLf
Response.Write "invoiceservicehdr affected: " & raUpdate & " | SQL: " & txtSQL & vbCrLf
Response.Write "-->" & vbCrLf

If hasError Then
    Response.Write "<script>alert('Gagal melakukan reset: " & Replace(errMsg, "'", "\'") & "'); window.close();</script>"
ElseIf raJurnalHdr = 0 Or raJurnalDtl = 0 Or raUpdate = 0 Then
    Response.Write "<script>alert('Peringatan: sebagian data tidak ditemukan (0 baris terupdate). Cek fcCode/fcba.'); window.close();</script>"
Else
    Response.Write "<script>"
    Response.Write "alert('Unposting berhasil: " & v_fcCode & "');"
    Response.Write "window.opener.location.href='http://localhost:92/show.asp?txtTable=InvoiceService&txtModule=Procurement';"
    Response.Write "window.close();"
    Response.Write "</script>"
End If
%>