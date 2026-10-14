<%@ Language=VBScript EnableSessionState=False %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include file="po_mail_sender.inc"-->
<%
'=====================================================================
' send_po_outbox.asp - memproses antrian email PO (PO_EMAIL_OUTBOX):
' membuat PDF ber-QR, mengirim email ke supplier, mencatat hasil.
'
' Dipanggil:
'  1. otomatis sesaat setelah approval final (po_mail_queue.inc -> PoMail_Kick)
'  2. berkala oleh Windows Task Scheduler lewat email/run_po_outbox.vbs (cadangan & percobaan ulang)
'
' Wajib menyertakan ?key=<PoMailCfg_ProcessKey>. Tidak memakai Session
' (EnableSessionState=False) agar tidak terkunci oleh halaman approval.
'=====================================================================
Response.Buffer = True
Response.ContentType = "text/plain"
Response.CharSet = "utf-8"
Server.ScriptTimeout = 600

If Request("key") = "" Or Request("key") <> PoMailCfg_ProcessKey Then
	Response.Status = "403 Forbidden"
	Response.Write "Akses ditolak."
	Response.End
End If

Dim maxRows
maxRows = 20
If IsNumeric(Request("max")) And Request("max") <> "" Then maxRows = CLng(Request("max"))
If maxRows < 1 Or maxRows > 100 Then maxRows = 20

Response.Write "[" & Now() & "] " & PoMail_ProcessOutbox(maxRows)
connStr.Close
Set connStr = Nothing
%>
