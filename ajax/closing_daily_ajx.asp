<!--#include file="../connDb1.Inc"-->
<!--#include file="../txtFunction.inc"-->
<%
dim txt_INDAY,txt_FCBA

txt_INDAY = Request("INDAY")
txt_FCBA = Request("FCBA")
txt_exec = Request("exec")
Dim txtajx
	
	if txt_exec = "doit" then
	  connstr.BeginTrans
		' CREATE EXEC CLOSING
		oQuery	= 	" BEGIN DAILYCLOSINGACTIVITY_API.CreateDailyClosing(" & _
					"	'" & txt_FCBA & "'" & _
					"	, TO_DATE('" & txt_INDAY & "','MM/DD/YYYY')" & _
					"	, '" & Session("USERNAME") & "'" & _
					"	, '" & Session("USERNAME") & "'" & _
					"	, '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
					" ); END; "
					
		'Response.Write oQuery & "<br/>"
		connStr.execute(oQuery)
		If err.number = 0 Then
			connStr.CommitTrans
		End If
		If err.number <> 0 Then
			connStr.RollbackTrans
		End If
		
	end if
	
	if txt_exec = "doupdate" then
	  connstr.BeginTrans
		' CREATE EXEC CLOSING
		oQuery	= 	" BEGIN DAILYCLOSINGACTIVITY_API.UpdateDailyOpening(" & _
					"	'" & txt_FCBA & "'" & _
					"	, TO_DATE('" & txt_INDAY & "','MM/DD/YYYY')" & _
					"	, '" & Session("USERNAME") & "'" & _
					"	, '" & Session("USERNAME") & "'" & _
					"	, '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
					" ); END; "
					
		'Response.Write oQuery & "<br/>"
		connStr.execute(oQuery)

		If err.number = 0 Then
			connStr.CommitTrans
		End If
		If err.number <> 0 Then
			connStr.RollbackTrans
		End If
		
	end if
txt_exec = ""
	
	 SQLata = "select DAILYCLOSINGACTIVITY_API.IsDailyClosing('" & txt_FCBA & "',TO_DATE('" & txt_INDAY & "','MM/DD/YYYY')) checker from dual"
	 'response.write(SQLata)
		Set rsQ = Server.CreateObject("ADODB.Recordset")
		rsQ.Open SQLata, connstr			   

		IF Not rsQ.eof then
			txtajx = rsQ("checker")
		else
			txtajx = "Not Found."
		End if
		
		rsQ.Close
		set rsQ = Nothing  
					  
	response.expires=-1
	
	
	'Output "no suggestion" if no hint were found
	'or output the correct values

	if UCase(txtajx) = "TRUE" then
		response.write("Closing")
	elseif UCase(txtajx) = "FALSE" then
		response.write("Not Yet Closing")
	End if
	
%>