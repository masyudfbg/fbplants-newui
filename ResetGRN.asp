<!--#include file="connDb1.Inc"-->
<%
Dim v_fcCode, v_fcba, txtSQL, bulansplit, tahunsplit

v_fcCode = Trim(Request("fcCode"))
v_fcba   = Trim(Request("fcba"))
v_date   = Request("tgl")
bulansplit  = left(v_date,2)
tglsplit    = mid(v_date,4,2)
thnsplit    = right(v_date,4)

' Validasi input
If Len(v_fcCode) = 0 Or Len(v_fcba) = 0 Then
    Response.Write "<script>alert('Parameter tidak valid!'); window.close();</script>"
    Response.End
End If

On Error Resume Next

delGLHdr = "delete from GLEDGERHDR WHERE DOCNO IN (SELECT DOCNO FROM GLEDGERDTL WHERE JOURNALNO = '"& v_fcCode &"')"
connStr.Execute(delGLHdr)

delGLDtl = "delete from GLEDGERDTL where JOURNALNO = '"& v_fcCode &"'"
connStr.Execute(delGLDtl)

delJurnalHdr = "delete from BKMJOURNALHDR where fcba = '"& v_fcba&"' and docno in ('"& v_fcCode &"','"& v_fcCode &"-INT')"
connStr.Execute(delJurnalHdr)

delJurnalDtl = "delete from BKMJOURNALDTL where fcba = '"& v_fcba&"' and docno in ('"& v_fcCode &"','"& v_fcCode &"-INT')"
connStr.Execute(delJurnalDtl)

txtSQL = "UPDATE receivenote SET rowstateid = 'Review Accounting', PostingFlg = 0 " & _
         "WHERE receivenotecode = '" & v_fcCode & "' " & _
         "AND fcba = '" & v_fcba & "'"

connStr.Execute(txtSQL)

qProcessPayroll = "BEGIN " & _ 
                "fac_closing_accounting_api.ProsesClosingMonth (" & thnsplit & "," & bulansplit & ",'" & v_fcba & "','" & Session("CodeName") & "'); " & _ 
                "END;"		
                
connstr.execute(qProcessPayroll)

' If Err.Number <> 0 Then
'     Response.Write "<script>alert('Gagal melakukan reset: " & Err.Description & "'); window.close();</script>"
' Else
'     Response.Write "<script>alert('Reset berhasil untuk GRN: " & v_fcCode & "'); window.opener.location.reload(); window.close();</script>"
' End If

On Error GoTo 0
%>