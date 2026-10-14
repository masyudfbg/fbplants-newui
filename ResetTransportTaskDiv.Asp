<!--#include file="connDb1.Inc"-->
<%
Dim v_OrderNo, v_fcba, txtSQL

v_OrderNo = Trim(Request("OrderNo"))

' Validasi input
If Len(v_OrderNo) = 0 Then
    Response.Write "<script>alert('Parameter tidak valid!'); window.close();</script>"
    Response.End
End If

On Error Resume Next

txtSQL = "UPDATE TRANSPORTTASKDIV_TAB SET ROWSTATE = 'Review Accounting', PostingFlg = 0 " & _    
         "WHERE orderno = '" & v_OrderNo & "' "

connStr.Execute(txtSQL)

If Err.Number <> 0 Then
    Response.Write "<script>alert('Gagal melakukan reset: " & Err.Description & "'); window.close();</script>"
Else
    Response.Write "<script>alert('Replanned berhasil untuk Order No: " & v_OrderNo & "'); window.opener.location.reload(); window.close();</script>"
End If

On Error GoTo 0
%>