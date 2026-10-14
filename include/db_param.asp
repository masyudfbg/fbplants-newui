<%
' =====================================================================
' Query berparameter (ADODB.Command) - mencegah SQL injection.
' Nilai dari user TIDAK disambung ke teks SQL; tulis "?" di SQL lalu kirim
' nilainya lewat array, sesuai urutan tanda "?".
'
'   Set rs = DbQuery("select fccode from users_tab where fccode = ? and fcpassword = ?", Array(user, md5(pwd)))
'   DbExec "update user_fcba_tab set isdefault = 'TRUE' where fcusercode = ? and fcba = ?", Array(user, ba)
'
' Butuh variabel koneksi connStr (connDb1.Inc).
' =====================================================================
Function DbCommand(sql, params)
	Dim cmd, i, v
	Set cmd = Server.CreateObject("ADODB.Command")
	Set cmd.ActiveConnection = connStr
	cmd.CommandText = sql
	cmd.CommandType = 1          ' adCmdText
	If IsArray(params) Then
		For i = 0 To UBound(params)
			v = params(i)
			If IsNull(v) Or IsEmpty(v) Then v = ""
			v = CStr(v)
			' adVarChar (200), adParamInput (1); ukuran minimal 1
			cmd.Parameters.Append cmd.CreateParameter("p" & i, 200, 1, Len(v) + 1, v)
		Next
	End If
	Set DbCommand = cmd
End Function

Function DbQuery(sql, params)
	Set DbQuery = DbCommand(sql, params).Execute
End Function

Sub DbExec(sql, params)
	DbCommand(sql, params).Execute , , 128   ' adExecuteNoRecords
End Sub
%>
