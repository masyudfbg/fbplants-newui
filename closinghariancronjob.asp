<%
Response.CharSet="euc-kr"
Session.codepage="949"
Response.codepage="949"
Response.ContentType="text/html;charset=euc-kr"
%>
<%
Response.LCID = 1046
Server.ScriptTimeout = 600
%>

<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->

<!--#includes file="include/JSON.asp" --> 
<!--#includes file="include/json.objectClass.asp"-->
<!--#includes file="include/aspJSON.asp" -->

<%
    Dim tahun, bulan
    tahun = Year(Now())
    bulan = Month(Now())

    Dim rsFCBA, sqlFCBA
    Dim successCount, failCount, skipCount, errorMsg
    Dim sqlLog
    successCount = 0
    failCount = 0
    skipCount = 0
    errorMsg = ""

    sqlFCBA = "SELECT fccode as fcba FROM businessunit WHERE isactive='TRUE' AND fctype='E'"
    Set rsFCBA = connstr.Execute(sqlFCBA)

    Do While Not rsFCBA.EOF
        Dim fcbaLoop, hasError
        fcbaLoop = rsFCBA("fcba")
        hasError = False

        On Error Resume Next

        ' --- WSI CLOSING ---
        Dim sqlclosing
        sqlclosing = "BEGIN " & _ 
                        "fac_journal_opr_wsi_api.OPR_WSI_CLOSING(" & tahun & "," & bulan & ",'" & fcbaLoop & "','CRONJOB'); " & _ 
                        "END;"      
        connstr.execute(sqlclosing)

        If Err.Number <> 0 Then
            failCount = failCount + 1
            hasError = True
            errorMsg = errorMsg & "&#x2716; [FCBA: <b>" & fcbaLoop & "</b>] WSI_CLOSING - " & Err.Description & "<br>"

            ' Simpan error WSI ke tabel log
            Dim errWSI
            errWSI = Replace(Err.Description, "'", "''")
            sqlLog = "INSERT INTO CRONJOB_CLOSING_LOG (tahun, bulan, fcba, proses, status, pesan, executed_by) " & _
                     "VALUES (" & tahun & "," & bulan & ",'" & fcbaLoop & "','WSI_CLOSING','ERROR','" & errWSI & "','CRONJOB')"
            connstr.execute(sqlLog)
            Err.Clear
        End If

        ' --- VRA CLOSING ---
        ' Tetap diproses meski WSI error
        sqlclosing = "BEGIN " & _ 
                        "fac_journal_opr_VRA_api.OPR_VRA_CLOSING(" & tahun & "," & bulan & ",'" & fcbaLoop & "','CRONJOB'); " & _ 
                        "END;"      
        connstr.execute(sqlclosing)

        If Err.Number <> 0 Then
            failCount = failCount + 1
            hasError = True
            errorMsg = errorMsg & "&#x2716; [FCBA: <b>" & fcbaLoop & "</b>] VRA_CLOSING - " & Err.Description & "<br>"

            ' Simpan error VRA ke tabel log
            Dim errVRA
            errVRA = Replace(Err.Description, "'", "''")
            sqlLog = "INSERT INTO CRONJOB_CLOSING_LOG (tahun, bulan, fcba, proses, status, pesan, executed_by) " & _
                     "VALUES (" & tahun & "," & bulan & ",'" & fcbaLoop & "','VRA_CLOSING','ERROR','" & errVRA & "','CRONJOB')"
            connstr.execute(sqlLog)
            Err.Clear
        End If

        On Error GoTo 0

        ' Hitung hasil per FCBA
        If hasError = False Then
            successCount = successCount + 1
        Else
            skipCount = skipCount + 1
        End If

        rsFCBA.MoveNext
    Loop

    rsFCBA.Close
    Set rsFCBA = Nothing

    ' --- Response Message ---
    If failCount = 0 Then
        Response.Write "<script>alert('Closing berhasil!\n\nPeriode: " & bulan & "/" & tahun & "\nTotal FC berhasil: " & successCount & "');</script>"
    ElseIf successCount = 0 Then
        Response.Write "<script>alert('Closing gagal!\n\nPeriode: " & bulan & "/" & tahun & "\nTotal FC gagal: " & skipCount & "');</script>"
    Else
        Response.Write "<script>alert('Closing selesai dengan sebagian error.\n\nPeriode: " & bulan & "/" & tahun & "\nFC Berhasil: " & successCount & "\nFC Ada Error: " & skipCount & "');</script>"
    End If

    If errorMsg <> "" Then
        Response.Write "<div style='font-family:Arial; font-size:12px; padding:10px; border:1px solid #e74c3c; background:#fdf3f3; border-radius:4px;'>"
        Response.Write "<b style='color:#e74c3c;'>&#x26A0; Detail Error (" & failCount & " error pada " & skipCount & " FC):</b><br><br>"
        Response.Write "<div style='color:#c0392b; line-height:1.8;'>" & errorMsg & "</div>"
        Response.Write "</div>"
    End If
%>