<!--#include file="po_mail_config.inc"-->
<%
'=====================================================================
' po_mail_queue.inc - antrian email PO yang sudah disetujui approver terakhir.
' Di-include oleh halaman approval (PurchaseOrderAuthorization.Asp,
' PurchaseOrderDetailsApproval.Asp). Ringan: hanya INSERT ke tabel
' PO_EMAIL_OUTBOX dan memicu pemroses. PDF dibuat & email dikirim oleh
' email/send_po_outbox.asp (dipicu segera + Task Scheduler sebagai cadangan).
'
' Kegagalan di sini TIDAK menggagalkan approval (error ditangkap & dicatat).
'=====================================================================
Dim PoMail_QueuedCount : PoMail_QueuedCount = 0

Function PoMail_Sql(s)
	PoMail_Sql = Replace("" & s, "'", "''")
End Function

' Masukkan PO ke antrian email. Dipanggil di cabang approval final,
' sebelum CommitTrans (bila ada) agar ikut batal saat approval di-rollback.
Sub PoMail_Queue(fcba, pocode)
	Dim q
	If Not PoMailCfg_Enabled Then Exit Sub
	If Len(Trim("" & pocode)) = 0 Or Len(Trim("" & fcba)) = 0 Then Exit Sub
	On Error Resume Next
	q = "INSERT INTO PO_EMAIL_OUTBOX (ID, FCBA, POCODE, STATUS, ATTEMPTS, FCENTRY, FCIP, CREATED_AT, UPDATED_AT)" & _
	    " VALUES (PO_EMAIL_OUTBOX_SEQ.NEXTVAL, '" & PoMail_Sql(fcba) & "', '" & PoMail_Sql(pocode) & "', 'PENDING', 0," & _
	    " '" & PoMail_Sql(Session("CodeName")) & "', '" & PoMail_Sql(Request.ServerVariables("REMOTE_ADDR")) & "', SYSDATE, SYSDATE)"
	connStr.Execute q
	If Err.Number = 0 Then
		PoMail_QueuedCount = PoMail_QueuedCount + 1
	Else
		Call WriteErrorLog("PO email queue gagal (" & fcba & " / " & pocode & "): " & Err.Description)
	End If
	Err.Clear
End Sub

' Picu pemroses antrian agar email terkirim segera (tunggu maks. +- 2 detik).
' Bila gagal / timeout, email tetap terkirim lewat jadwal Task Scheduler.
Sub PoMail_Kick()
	Dim http
	If PoMail_QueuedCount = 0 Then Exit Sub
	On Error Resume Next
	Set http = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
	http.setTimeouts 2000, 2000, 2000, 2000
	http.open "GET", PoMailCfg_BaseUrl & "/email/send_po_outbox.asp?key=" & Server.URLEncode(PoMailCfg_ProcessKey), False
	http.send
	Set http = Nothing
	Err.Clear
End Sub
%>
