' =====================================================================
' run_po_outbox.vbs - dijalankan Windows Task Scheduler (mis. tiap 5 menit)
' untuk memproses antrian email PO yang belum terkirim / perlu dicoba ulang.
' URL & kunci dibaca otomatis dari po_mail_config.inc di folder yang sama
' (PoMailCfg_BaseUrl & PoMailCfg_ProcessKey).
'
' Contoh tugas terjadwal (Command Prompt sebagai Administrator; sesuaikan path):
'   schtasks /Create /TN "FB-Plants PO Email" /SC MINUTE /MO 5 /RU SYSTEM ^
'     /TR "cscript.exe //nologo D:\WebSites\fb-plants\email\run_po_outbox.vbs"
' =====================================================================
Option Explicit
Dim fso, cfgPath, cfg, baseUrl, key, http

Set fso = CreateObject("Scripting.FileSystemObject")
cfgPath = fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "po_mail_config.inc")
If Not fso.FileExists(cfgPath) Then
	WScript.Echo "Tidak menemukan " & cfgPath
	WScript.Quit 1
End If
cfg = fso.OpenTextFile(cfgPath, 1).ReadAll
baseUrl = ConfigValue(cfg, "PoMailCfg_BaseUrl")
key = ConfigValue(cfg, "PoMailCfg_ProcessKey")
If baseUrl = "" Or key = "" Then
	WScript.Echo "PoMailCfg_BaseUrl / PoMailCfg_ProcessKey belum diisi di po_mail_config.inc"
	WScript.Quit 1
End If

On Error Resume Next
Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
http.setTimeouts 5000, 5000, 30000, 600000
http.open "GET", baseUrl & "/email/send_po_outbox.asp?key=" & key, False
http.send
If Err.Number <> 0 Then
	WScript.Echo "Gagal memanggil pemroses: " & Err.Description
	WScript.Quit 1
End If
WScript.Echo http.Status & " " & http.responseText
If http.Status <> 200 Then WScript.Quit 1

Function ConfigValue(text, name)
	Dim re, m
	Set re = New RegExp
	re.IgnoreCase = True
	re.Pattern = "^\s*" & name & "\s*=\s*""([^""]*)"""
	re.Multiline = True
	Set m = re.Execute(text)
	If m.Count > 0 Then ConfigValue = m(0).SubMatches(0) Else ConfigValue = ""
End Function
