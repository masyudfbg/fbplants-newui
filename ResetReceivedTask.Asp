<!--#include file="connDb1.Inc"-->
<%
Dim v_fcCode, v_fcba, txtSQL, bulansplit, tahunsplit

v_fcCode = Trim(Request("fcCode"))
v_OrderNo = Trim(Request("OrderNO"))
v_fcba   = Trim(Request("fcba"))
v_date   = Request("tgl")
bulansplit  = left(v_date,2)
tglsplit    = mid(v_date,4,2)
thnsplit    = right(v_date,4)

' Validasi input
If Len(v_OrderNo) = 0 Or Len(v_fcba) = 0 Then
    Response.Write "<script>alert('Parameter tidak valid!'); window.close();</script>"
    Response.End
End If

On Error Resume Next

txtSQL = "UPDATE RECEIVEDTASK_TAB SET rowstate = 'Review Accounting', PostingFlg = 0 " & _
         "WHERE orderno = '" & v_OrderNo & "' " & _
         "AND fcba = '" & v_fcba & "'"

connStr.Execute(txtSQL)


delGLHdr = "delete from GLEDGERHDR WHERE DOCNO IN (SELECT DOCNO FROM GLEDGERDTL WHERE JOURNALNO = '"& v_fcCode &"')"
connStr.Execute(delGLHdr)

delGLDtl = "delete from GLEDGERDTL where JOURNALNO = '"& v_fcCode &"'"
connStr.Execute(delGLDtl)

delJurnalHdr = "delete from BKMJOURNALHDR where fcba = '"& v_fcba&"' and docno in ('"& v_fcCode &"','"& v_fcCode &"-INT')"
connStr.Execute(delJurnalHdr)

delJurnalDtl = "delete from BKMJOURNALDTL where fcba = '"& v_fcba&"' and docno in ('"& v_fcCode &"','"& v_fcCode &"-INT')"
connStr.Execute(delJurnalDtl)

If Err.Number <> 0 Then
    Response.Write "<script>alert('Gagal melakukan reset header: " & Err.Description & "'); window.close();</script>"
    Response.End
End If

txtSQL2 = "UPDATE RECEIVEDTASKDTL_TAB SET quantityreceived = 0 " & _
          "WHERE orderno = '" & v_OrderNo & "' " & _
          "AND fcba = '" & v_fcba & "'"

connStr.Execute(txtSQL2)

qProcessPayroll = "BEGIN " & _ 
                "fac_closing_accounting_api.ProsesClosingMonth (" & thnsplit & "," & bulansplit & ",'" & v_fcba & "','" & Session("CodeName") & "'); " & _ 
                "END;"		
connstr.execute(qProcessPayroll)

If Err.Number <> 0 Then
    Response.Write "<script>alert('Gagal melakukan reset detail: " & Err.Description & "'); window.close();</script>"
Else
    Response.Write "<script>alert('Replanned berhasil untuk Order No: " & v_OrderNo & "'); window.opener.location.reload(); window.close();</script>"
End If

On Error GoTo 0
%>